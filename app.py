"""
pptx-reviewer Web UI
Flask バックエンド。PPTXをアップロードしてAIによるレビューレポートを生成する。

対応AIプロバイダー（.env の AI_PROVIDER で切り替え）:
  AI_PROVIDER=anthropic  → Claude (Anthropic API) ※デフォルト
  AI_PROVIDER=openai     → OpenAI / Azure OpenAI / LiteLLM Proxy

OpenAI 系の接続先切り替え（OPENAI_API_TYPE で指定）:
  OPENAI_API_TYPE=openai  → 標準 OpenAI API（デフォルト）
  OPENAI_API_TYPE=azure   → Azure OpenAI Service

LiteLLM Proxy など互換エンドポイントの場合:
  OPENAI_API_TYPE=openai + OPENAI_BASE_URL=https://xxxxx.net/

起動方法:
  cd web
  python app.py
"""
import csv
import io
import json
import logging
import logging.handlers
import os
import shutil
import socket
import subprocess
import sys
import tempfile
import threading
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime
from pathlib import Path

import anthropic
import requests as requests_lib
import yaml
from dotenv import load_dotenv
from flask import Flask, jsonify, render_template, request

try:
    from openai import OpenAI as OpenAIClient
    _OPENAI_AVAILABLE = True
except ImportError:
    _OPENAI_AVAILABLE = False

# ---------------------------------------------------------------------------
# 設定ファイル読み込み・ログ設定
# ---------------------------------------------------------------------------
load_dotenv()

BASE_DIR = Path(__file__).parent


def _load_config() -> dict:
    """config.yaml を読み込む。ファイルがなければ空辞書を返す。"""
    config_path = BASE_DIR / "config.yaml"
    if config_path.exists():
        with open(config_path, encoding="utf-8") as f:
            return yaml.safe_load(f) or {}
    return {}


_config = _load_config()
_log_cfg = _config.get("logging", {})
_LOG_RETENTION_DAYS = int(os.getenv("LOG_RETENTION_DAYS", str(_log_cfg.get("retention_days", 7))))
_LOG_FORMAT = '{"timestamp": "%(asctime)s", "level": "%(levelname)s", "message": "%(message)s"}'
_LOG_DIR = BASE_DIR / _log_cfg.get("log_dir", "logs")
_LOG_DIR.mkdir(exist_ok=True)

_file_handler = logging.handlers.TimedRotatingFileHandler(
    _LOG_DIR / "app.log",
    when="midnight",
    backupCount=_LOG_RETENTION_DAYS,
    encoding="utf-8",
)
_file_handler.setFormatter(logging.Formatter(_LOG_FORMAT))

logging.basicConfig(
    level=logging.INFO,
    format=_LOG_FORMAT,
    handlers=[logging.StreamHandler(), _file_handler],
)
logger = logging.getLogger(__name__)


def _resolve_skill_dir() -> Path:
    """スキルフォルダのパスを解決する。

    優先順位:
    1. 環境変数 SKILL_DIR（明示指定）
    2. Codex CLI の標準パス  : %USERPROFILE%\\.codex\\skills\\pptx-review-text
    3. Claude Code の標準パス: %USERPROFILE%\\.claude\\skills\\pptx-reviewer
    4. このファイルの相対パス : ../skills/pptx-reviewer（開発時のデフォルト）
    """
    # 1. 環境変数
    env_val = os.getenv("SKILL_DIR", "").strip()
    if env_val:
        p = Path(env_val)
        if p.exists():
            logger.info("SKILL_DIR (env): %s", p)
            return p
        logger.warning("SKILL_DIR env で指定されたパスが見つかりません: %s", p)

    home = Path.home()

    # 2. Codex CLI 標準パス
    codex_path = home / ".codex" / "skills" / "pptx-review-text"
    if codex_path.exists():
        logger.info("SKILL_DIR (Codex CLI): %s", codex_path)
        return codex_path

    # 3. Claude Code 標準パス
    claude_path = home / ".claude" / "skills" / "pptx-reviewer"
    if claude_path.exists():
        logger.info("SKILL_DIR (Claude Code): %s", claude_path)
        return claude_path

    # 4. 相対パス（開発時フォールバック）
    relative = BASE_DIR.parent / "skills" / "pptx-reviewer"
    logger.info("SKILL_DIR (relative fallback): %s", relative)
    return relative


SKILL_DIR = _resolve_skill_dir()
EXTRACT_SCRIPT = SKILL_DIR / "scripts" / "extract_pptx.py"
TERMINOLOGY_SCRIPT = SKILL_DIR / "scripts" / "check_terminology.py"
TERMINOLOGY_PATH = SKILL_DIR / "references" / "terminology.json"
PERSPECTIVES_PATH = SKILL_DIR / "references" / "custom_perspectives.json"
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)


def _cleanup_uploads_on_startup() -> None:
    """起動時に uploads/ 直下の残留ファイル・ディレクトリをクリーンアップする。"""
    for entry in UPLOAD_DIR.iterdir():
        if entry.is_file() and entry.suffix == ".md":
            try:
                entry.unlink()
                logger.info("起動時クリーンアップ(.md): %s", entry.name)
            except OSError as e:
                logger.warning("起動時クリーンアップ失敗: %s / %s", entry.name, e)
        elif entry.is_dir():
            shutil.rmtree(entry, ignore_errors=True)
            logger.info("起動時クリーンアップ(dir): %s", entry.name)


_cleanup_uploads_on_startup()

MAX_UPLOAD_MB = 50

# チャンクサイズ: config.yaml → 環境変数 REVIEW_CHUNK_SIZE → デフォルト 3
_review_cfg = _config.get("review", {})
_REVIEW_CHUNK_SIZE = int(os.getenv("REVIEW_CHUNK_SIZE", str(_review_cfg.get("chunk_size", 3))))

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024


# ---------------------------------------------------------------------------
# レビュープロンプト生成
# ---------------------------------------------------------------------------
def _build_perspectives_section(selected_perspectives: list[dict] | None) -> str:
    """選択された固有観点からプロンプトの観点3セクションを生成する。"""
    if not selected_perspectives:
        return ""

    lines = ["### 観点3：固有観点", "以下のユーザー定義観点に基づいてレビューを行う。\n"]
    for cat in selected_perspectives:
        category = cat.get("category", "")
        items = cat.get("items", [])
        if not items:
            continue
        lines.append(f"#### {category}")
        for item in items:
            perspective = item.get("perspective", "")
            notes = item.get("notes", "")
            if notes:
                lines.append(f"- {perspective}: {notes}")
            else:
                lines.append(f"- {perspective}")
        lines.append("")
    return "\n".join(lines)


def _build_perspectives_count_row(selected_perspectives: list[dict] | None) -> str:
    """固有観点のカテゴリ別件数行を生成する。"""
    if not selected_perspectives:
        return ""
    return "| 🔍 固有観点 | N件 |"


def build_review_prompt(
    filename: str,
    extract_data: dict,
    terminology_data: dict,
    selected_perspectives: list[dict] | None = None,
) -> str:
    """SKILL.md の Step 5–6 に相当するレビュー指示プロンプトを生成する。"""
    today = date.today().strftime("%Y-%m-%d")
    total_slides = extract_data.get("total_slides", "?")
    reviewed_slides = extract_data.get("reviewed_slides", [])
    page_label = (
        "全スライド"
        if len(reviewed_slides) == total_slides
        else ", ".join(str(p) for p in reviewed_slides) + " ページ"
    )

    extract_json_str = json.dumps(extract_data, ensure_ascii=False, indent=2)
    terminology_json_str = json.dumps(terminology_data, ensure_ascii=False, indent=2)

    # 観点3: 固有観点セクション（選択された場合のみ）
    perspectives_section = _build_perspectives_section(selected_perspectives)
    perspectives_count_row = _build_perspectives_count_row(selected_perspectives)

    # 観点数の決定
    has_perspectives = bool(selected_perspectives)
    perspective_count = "4" if has_perspectives else "3"
    perspectives_emoji_line = "🎯 = 固有観点（観点3）\n" if has_perspectives else ""

    return f"""あなたはPowerPoint資料の品質レビュアーです。
以下のスライド抽出データと用語チェック結果をもとに、{perspective_count}つの観点でレビューを行い、
指定のMarkdownフォーマットでレポートを出力してください。

## レビュー観点

### 観点1：文章校正・表記ゆれ
- 用語チェック結果（terminology_data）の誤表記を漏れなく報告する
- リスト外の表記ゆれも指摘する（同じ概念に複数の表記が混在）
- 誤字・脱字・文法エラー（助詞の誤り、読点の欠落など）
- 英単語・英語の関数名・技術用語は指摘対象外
- notes（ノートペイン）のテキストも対象（箇所は「ノートペイン」と明記）
- **以下は指摘禁止**（スライド資料として正常な書き方）:
  「見出し的項目が連続して区切りが不明確」「文が断片的で読み取りにくい」「文章として不完全」
  → 箇条書き・体言止め・短文はスライドの標準書式であり指摘対象外

### 観点2：論理的整合性
- スライド間の矛盾（前後で事実が食い違っていないか）
- 主語・述語のねじれ
- 因果関係の破綻（接続詞と内容が対応していないケース）
- 複数スライドにまたがる矛盾: 主要箇所を指摘箇所に引用し、参照先は指摘事項内に（スライドNの「〜」と矛盾）として記述する

{perspectives_section}### 観点{perspective_count}：文体の調整
- です・ます調とだ・である調の混在
- 「など」「等」「場合によっては」の多用
- 「全体的に文体が乱れている」などの総論的指摘は不可。必ず「どのテキストのどの語句が混在しているか」を特定すること

## カテゴリ絵文字（テーブルのカテゴリ列は絵文字のみ使用すること）
📝 = 文章校正・表記ゆれ（観点1）
🔗 = 論理的整合性（観点2）
{perspectives_emoji_line}👥 = 文体（観点{perspective_count}）

## 出力フォーマット（厳守）

以下のMarkdownフォーマットを**厳密に**守ること。セクション順序・表の列構成を変えてはならない。

```
# PowerPoint レビューレポート

**ファイル**: {filename}
**レビュー日時**: {today}
**総スライド数**: {total_slides} 枚（ファイル全体）
**レビュー対象ページ**: {page_label}
**対象読者**: 顧客向け

---

## カテゴリ別 指摘件数

| カテゴリ | 件数 |
|---------|------|
| 📝 表記ゆれ・文章校正 | N件 |
| 🔗 論理的整合性 | N件 |
{perspectives_count_row}
| 👥 文体 | N件 |
| **合計** | **N件** |

---

## スライド別 指摘事項

### スライド N：タイトル

| カテゴリ | 指摘事項 | 指摘箇所 |
|---------|---------|---------|
| 📝 | 用語統一リスト（正: サーバー）と異なる表記が使用されている | タイトル「サーバ構成」→「サーバ」 |
| 👥 | です/ます調の中に「〜だ」形が混在している | テキストボックス「処理は完了だ」→「だ」 |

（指摘がないスライドは省略）
```

## 出力ルール
- 指摘箇所: 問題のあるテキストを**直接引用**し、場所（shape_kind: タイトル／テキストボックス／図形／表／グラフ／ノートペイン）も記載する（合計50文字以内）
  ✅ 良い例: タイトル「サーバ構成」→「サーバ」
  ✅ 良い例: テキストボックス「処理は完了だ」→「だ」
  ❌ 悪い例: テキストボックス 2行目（どの語句が問題か不明）
- 指摘事項: 「何が問題か」を断言する一文で記述する
  ✅ 良い例: 「用語統一リスト（正: サーバー）と異なる表記が使用されている」
  ❌ 悪い例: 「表記ゆれが見受けられます」（曖昧）
- カテゴリ別件数は正確に集計してから記載
- コードブロック（```）は出力に含めない。Markdownをそのまま出力する

---

## スライド抽出データ

```json
{extract_json_str}
```

## 用語チェック結果

```json
{terminology_json_str}
```
"""


# ---------------------------------------------------------------------------
# ページ指定のユーティリティ
# ---------------------------------------------------------------------------
def expand_page_ranges(pages_str: str) -> str:
    """ページ指定文字列のチルダ範囲表記を展開してカンマ区切り文字列に変換する。

    例:
      "1～5,10" → "1,2,3,4,5,10"
      "3~5,11"  → "3,4,5,11"
      "1,3,5"   → "1,3,5"（変更なし）
    """
    result: list[int] = []
    seen: set[int] = set()
    for token in pages_str.split(","):
        token = token.strip()
        if not token:
            continue
        token_normalized = token.replace("～", "~")
        if "~" in token_normalized:
            parts = token_normalized.split("~", 1)
            try:
                start = int(parts[0].strip())
                end = int(parts[1].strip())
                if start < 1 or end < 1 or start > end:
                    logger.warning("無効なページ範囲のため無視: %s", token)
                    continue
                for n in range(start, end + 1):
                    if n not in seen:
                        result.append(n)
                        seen.add(n)
            except ValueError:
                logger.warning("ページ範囲のパースに失敗したため無視: %s", token)
        else:
            try:
                n = int(token)
                if n < 1:
                    logger.warning("ページ番号は1以上で指定してください（無視: %s）", token)
                    continue
                if n not in seen:
                    result.append(n)
                    seen.add(n)
            except ValueError:
                logger.warning("無効なページ番号のため無視: %s", token)
    return ",".join(str(n) for n in result)


# ---------------------------------------------------------------------------
# スクリプト実行ヘルパー
# ---------------------------------------------------------------------------
def run_extract(pptx_path: str, pages: str | None) -> dict:
    """extract_pptx.py を実行してJSONを返す。"""
    cmd = [sys.executable, str(EXTRACT_SCRIPT), pptx_path]
    if pages:
        cmd.extend(["--pages", pages])
    logger.info("extract_pptx.py 実行: %s", " ".join(cmd))
    env = {**os.environ, "PYTHONUTF8": "1"}
    result = subprocess.run(
        cmd, capture_output=True, text=True, encoding="utf-8", env=env
    )
    if result.returncode != 0:
        raise RuntimeError(f"テキスト抽出エラー:\n{result.stderr}")
    return json.loads(result.stdout)


def run_terminology_check(extract_stdout: str) -> dict:
    """check_terminology.py を実行してJSONを返す。"""
    cmd = [sys.executable, str(TERMINOLOGY_SCRIPT), "-"]
    logger.info("check_terminology.py 実行")
    env = {**os.environ, "PYTHONUTF8": "1"}
    result = subprocess.run(
        cmd,
        input=extract_stdout,
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=env,
    )
    if result.returncode != 0:
        raise RuntimeError(f"用語チェックエラー:\n{result.stderr}")
    return json.loads(result.stdout)


# ---------------------------------------------------------------------------
# Teams 通知
# ---------------------------------------------------------------------------
def _notify_teams_api_error(error_message: str) -> None:
    """API 認証エラー（401）発生時に Teams へ通知する。

    webhook_type:
      incoming_webhook : Teams Incoming Webhook（MessageCard 形式）
      power_automate   : Power Automate「HTTP 要求を受信したとき」トリガー（Adaptive Card 形式）
    """
    teams_cfg = _config.get("teams", {})
    if not teams_cfg.get("enabled"):
        return
    webhook_url = teams_cfg.get("webhook_url", "").strip()
    if not webhook_url:
        logger.warning("Teams 通知: webhook_url が config.yaml に設定されていません")
        return

    import datetime as _dt

    webhook_type = teams_cfg.get("webhook_type", "power_automate").strip().lower()
    now_str = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    host = socket.gethostname()
    detail = error_message[:500]

    if webhook_type == "incoming_webhook":
        # 旧来の Incoming Webhook（MessageCard 形式）
        payload: dict = {
            "@type": "MessageCard",
            "@context": "http://schema.org/extensions",
            "themeColor": "FF0000",
            "summary": "API 認証エラー (401)",
            "sections": [
                {
                    "activityTitle": "pptx-reviewer: API 接続エラー (Error code: 401)",
                    "activitySubtitle": "APIキーの認証に失敗しました。.env ファイルのAPIキーを確認してください。",
                    "facts": [
                        {"name": "ホスト", "value": host},
                        {"name": "発生時刻", "value": now_str},
                        {"name": "エラー詳細", "value": detail},
                    ],
                }
            ],
        }
    else:
        # Power Automate「HTTP 要求を受信したとき」→ Teams チャネルへ投稿
        # Adaptive Card 形式（Teams クライアントで直接レンダリング）
        payload = {
            "type": "message",
            "attachments": [
                {
                    "contentType": "application/vnd.microsoft.card.adaptive",
                    "content": {
                        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                        "type": "AdaptiveCard",
                        "version": "1.4",
                        "body": [
                            {
                                "type": "TextBlock",
                                "text": "pptx-reviewer: API 認証エラー (401)",
                                "weight": "Bolder",
                                "size": "Medium",
                                "color": "Attention",
                            },
                            {
                                "type": "FactSet",
                                "facts": [
                                    {"title": "ホスト", "value": host},
                                    {"title": "発生時刻", "value": now_str},
                                    {"title": "エラー詳細", "value": detail},
                                ],
                            },
                            {
                                "type": "TextBlock",
                                "text": "APIキーの認証に失敗しました。管理者に連絡してAPIキーの再発行・登録を依頼してください。",
                                "wrap": True,
                                "color": "Attention",
                            },
                        ],
                    },
                }
            ],
        }

    try:
        resp = requests_lib.post(webhook_url, json=payload, timeout=10)
        resp.raise_for_status()
        logger.info("Teams 通知送信完了 (type=%s)", webhook_type)
    except requests_lib.RequestException as e:
        logger.warning("Teams 通知失敗: %s", e)


# ---------------------------------------------------------------------------
# AI API 呼び出し（Anthropic / OpenAI 切り替え対応）
# ---------------------------------------------------------------------------
def call_claude_review(prompt: str) -> str:
    """Anthropic Claude API を呼び出してレビューレポートを返す。"""
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError(
            "ANTHROPIC_API_KEY が設定されていません。web/.env ファイルを確認してください。"
        )
    client = anthropic.Anthropic(api_key=api_key)
    logger.info("Claude API (Anthropic) 呼び出し中...")
    try:
        message = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=8192,
            messages=[{"role": "user", "content": prompt}],
        )
    except anthropic.AuthenticationError as e:
        err_str = str(e)
        logger.error("Claude API 認証エラー (401): %s", err_str)
        _notify_teams_api_error(err_str)
        raise RuntimeError(
            "API_AUTH_ERROR: APIキーの認証に失敗しました。"
            "管理者に連絡してAPIキーの再発行・登録を依頼してください。"
            "（画面上部の「OPENAI KEY SETTING」ボタンから登録可能です）"
        ) from e
    text = message.content[0].text if message.content else ""
    if not text or not text.strip():
        logger.warning("Claude API が空のレスポンスを返しました")
        raise RuntimeError("AIレビューが空のレスポンスを返しました。再度お試しください。")
    return text


def call_openai_review(prompt: str) -> str:
    """OpenAI API / Azure OpenAI / LiteLLM Proxy を呼び出してレビューレポートを返す。

    OPENAI_API_TYPE=openai (default):
        標準 OpenAI API。OPENAI_BASE_URL を設定すると LiteLLM 等の互換プロキシも使用可能。
    OPENAI_API_TYPE=azure:
        Azure OpenAI Service。AZURE_OPENAI_ENDPOINT と AZURE_OPENAI_API_VERSION が必要。
        OPENAI_MODEL にはデプロイメント名を指定する。
    """
    if not _OPENAI_AVAILABLE:
        raise RuntimeError(
            "openai パッケージがインストールされていません。"
            "`pip install openai` を実行してください。"
        )
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError(
            "OPENAI_API_KEY が設定されていません。web/.env ファイルを確認してください。"
        )

    api_type = os.getenv("OPENAI_API_TYPE", "openai").strip().lower()
    model = os.getenv("OPENAI_MODEL", "gpt-4o")

    if api_type == "azure":
        from openai import AzureOpenAI as AzureOpenAIClient
        endpoint = os.getenv("AZURE_OPENAI_ENDPOINT", "").strip()
        if not endpoint:
            raise RuntimeError(
                "OPENAI_API_TYPE=azure のとき AZURE_OPENAI_ENDPOINT の設定が必要です。"
            )
        api_version = os.getenv("AZURE_OPENAI_API_VERSION", "2024-02-01")
        client: OpenAIClient = AzureOpenAIClient(  # type: ignore[assignment]
            api_key=api_key,
            azure_endpoint=endpoint,
            api_version=api_version,
        )
        logger.info(
            "Azure OpenAI API 呼び出し中... (endpoint=%s, deployment=%s, api_version=%s)",
            endpoint, model, api_version,
        )
    else:
        # 標準 OpenAI API または LiteLLM 互換プロキシ
        base_url = os.getenv("OPENAI_BASE_URL", "").strip() or None
        client = OpenAIClient(api_key=api_key, base_url=base_url)
        if base_url:
            logger.info(
                "OpenAI互換API (LiteLLM等) 呼び出し中... (base_url=%s, model=%s)",
                base_url, model,
            )
        else:
            logger.info("OpenAI API 呼び出し中... (model=%s)", model)

    try:
        response = client.chat.completions.create(
            model=model,
            max_tokens=8192,
            messages=[{"role": "user", "content": prompt}],
        )
    except Exception as e:
        # openai.AuthenticationError は _OPENAI_AVAILABLE=False のときインポートできないため
        # 文字列で 401 を検出する
        err_str = str(e)
        if "401" in err_str:
            logger.error("OpenAI API 認証エラー (401): %s", err_str)
            _notify_teams_api_error(err_str)
            raise RuntimeError(
                "API_AUTH_ERROR: APIキーの認証に失敗しました。"
                "管理者に連絡してAPIキーの再発行・登録を依頼してください。"
                "（画面上部の「OPENAI KEY SETTING」ボタンから登録可能です）"
            ) from e
        raise
    content = response.choices[0].message.content
    if content is None or not content.strip():
        logger.warning(
            "OpenAI API が空のレスポンスを返しました (finish_reason=%s)",
            response.choices[0].finish_reason,
        )
        raise RuntimeError("AIレビューが空のレスポンスを返しました。再度お試しください。")
    return content


def detect_provider() -> str:
    """使用する AI プロバイダーを決定する。

    優先順位:
    1. AI_PROVIDER 環境変数が明示されていればそれを使う
    2. 未設定の場合、利用可能な API キーから自動検出する
       - OPENAI_API_KEY のみ設定済み  → openai
       - ANTHROPIC_API_KEY が設定済み  → anthropic（デフォルト）
    """
    explicit = os.getenv("AI_PROVIDER", "").strip().lower()
    if explicit:
        return explicit
    if os.getenv("OPENAI_API_KEY") and not os.getenv("ANTHROPIC_API_KEY"):
        logger.info("AI_PROVIDER 未設定: OPENAI_API_KEY を検出したため openai を使用します")
        return "openai"
    return "anthropic"


_TEST_MODE_RESPONSE = """\
# PowerPoint レビューレポート

**ファイル**: dummy_proposal.pptx
**レビュー日時**: 2026-04-03
**総スライド数**: 5 枚（ファイル全体）
**レビュー対象ページ**: 全スライド
**対象読者**: 顧客向け

---

## カテゴリ別 指摘件数

| カテゴリ | 件数 |
|---------|------|
| 📝 表記ゆれ・文章校正 | 3件 |
| 🔗 論理的整合性 | 2件 |
| 👥 文体 | 2件 |
| **合計** | **7件** |

---

## スライド別 指摘事項

### スライド 2：システム概要

| カテゴリ | 指摘事項 | 指摘箇所 |
|---------|---------|---------|
| 📝 | 用語統一リスト（正: サーバ）と異なる表記が使用されている | テキストボックス「サーバー構成」→「サーバー」 |
| 👥 | 「RBAC」が説明なしに使用されている | テキストボックス「RBACを適用」→「RBAC」 |

### スライド 3：課題と対策

| カテゴリ | 指摘事項 | 指摘箇所 |
|---------|---------|---------|
| 📝 | 用語統一リスト（正: インタフェース）と異なる表記が使用されている | テキストボックス「インターフェース設計」→「インターフェース」 |
| 🔗 | 主語「課題は」に対して述語「改善します」が意味的にねじれている | テキストボックス「課題は改善します」→「課題は」 |
| 🔗 | 文字数約612文字。1スライドの情報量として過多 | テキストボックス（全体・612文字） |

### スライド 5：まとめ

| カテゴリ | 指摘事項 | 指摘箇所 |
|---------|---------|---------|
| 📝 | 「データベース」と「DB」が同一スライド内で混在している | テキストボックス「DBへの接続」→「DB」 |
| 👥 | 「です・ます調」の中に「〜だ」形が混在している | テキストボックス「処理は完了だ」→「だ」 |
"""


def call_ai_review(prompt: str) -> str:
    """検出したプロバイダーに応じて Anthropic または OpenAI を呼び出す。空レスポンス時は1回リトライ。

    TEST_MODE=true の場合は API 呼び出しをスキップし、固定のモックレスポンスを返す。
    """
    if os.getenv("TEST_MODE", "").strip().lower() == "true":
        logger.info("TEST_MODE 有効: モックレスポンスを返します")
        return _TEST_MODE_RESPONSE

    provider = detect_provider()
    logger.info("AI プロバイダー: %s", provider)
    call_fn = call_openai_review if provider == "openai" else call_claude_review

    max_attempts = 2
    for attempt in range(1, max_attempts + 1):
        try:
            return call_fn(prompt)
        except RuntimeError as e:
            if "空のレスポンス" in str(e) and attempt < max_attempts:
                logger.warning(
                    "AI レスポンスが空のため再試行します (attempt=%d/%d)",
                    attempt, max_attempts,
                )
                continue
            raise


# ---------------------------------------------------------------------------
# チャンク分割レビュー
# ---------------------------------------------------------------------------

def _split_extract_data(extract_data: dict, chunk_size: int) -> list[dict]:
    """スライドデータを chunk_size 枚ずつのチャンクに分割する。"""
    slides = extract_data.get("slides", [])
    total = extract_data.get("total_slides", len(slides))
    chunks: list[dict] = []
    for i in range(0, len(slides), chunk_size):
        chunk_slides = slides[i : i + chunk_size]
        chunks.append({
            "total_slides": total,
            "reviewed_slides": [s["slide_number"] for s in chunk_slides],
            "slides": chunk_slides,
        })
    return chunks


def _build_chunk_prompt(
    filename: str,
    chunk_data: dict,
    terminology_data: dict,
    selected_perspectives: list[dict] | None,
    chunk_index: int,
    total_chunks: int,
) -> str:
    """1チャンク分のスライド別指摘事項のみを返すプロンプトを生成する。"""
    slides = chunk_data.get("reviewed_slides", [])
    slide_range = f"スライド {slides[0]}〜{slides[-1]}" if slides else "不明"
    extract_json_str = json.dumps(chunk_data, ensure_ascii=False, indent=2)
    terminology_json_str = json.dumps(terminology_data, ensure_ascii=False, indent=2)

    perspectives_section = _build_perspectives_section(selected_perspectives)
    has_perspectives = bool(selected_perspectives)
    perspective_count = "4" if has_perspectives else "3"
    perspectives_emoji_line = "🎯 = 固有観点（観点3）\n" if has_perspectives else ""

    return f"""あなたはPowerPoint資料の品質レビュアーです。
以下は「{filename}」の一部（チャンク {chunk_index}/{total_chunks}、{slide_range}）です。
{perspective_count}つの観点でレビューし、**スライド別の指摘事項のみ**を出力してください。
件数表は出力不要です。

## レビュー観点

### 観点1：文章校正・表記ゆれ
- 用語チェック結果（terminology_data）の誤表記を漏れなく報告する
- リスト外の表記ゆれも指摘する（同じ概念に複数の表記が混在）
- 誤字・脱字・文法エラー（助詞の誤り、読点の欠落など）
- 英単語・英語の関数名・技術用語は指摘対象外
- notes（ノートペイン）のテキストも対象（箇所は「ノートペイン」と明記）
- **以下は指摘禁止**（スライド資料として正常な書き方）:
  「見出し的項目が連続して区切りが不明確」「文が断片的で読み取りにくい」「文章として不完全」
  → 箇条書き・体言止め・短文はスライドの標準書式であり指摘対象外

### 観点2：論理的整合性
- スライド間の矛盾（前後で事実が食い違っていないか）
- 主語・述語のねじれ
- 因果関係の破綻（接続詞と内容が対応していないケース）
- 複数スライドにまたがる矛盾: 主要箇所を指摘箇所に引用し、参照先は指摘事項内に（スライドNの「〜」と矛盾）として記述する

{perspectives_section}### 観点{perspective_count}：文体の調整
- です・ます調とだ・である調の混在
- 「など」「等」「場合によっては」の多用
- 「全体的に文体が乱れている」などの総論的指摘は不可。必ず「どのテキストのどの語句が混在しているか」を特定すること

## カテゴリ絵文字（テーブルのカテゴリ列は絵文字のみ使用すること）
📝 = 文章校正・表記ゆれ（観点1）
🔗 = 論理的整合性（観点2）
{perspectives_emoji_line}👥 = 文体（観点{perspective_count}）

## 出力フォーマット（厳守）

指摘があるスライドのみ、以下の形式で出力してください。指摘が全くない場合は「指摘なし」とのみ出力してください。

### スライド N：タイトル

| カテゴリ | 指摘事項 | 指摘箇所 |
|---------|---------|---------|
| 📝 | 用語統一リスト（正: サーバー）と異なる表記が使用されている | タイトル「サーバ構成」→「サーバ」 |
| 👥 | です/ます調の中に「〜だ」形が混在している | テキストボックス「処理は完了だ」→「だ」 |

## 出力ルール
- 指摘箇所: 問題のあるテキストを**直接引用**し、場所（shape_kind: タイトル／テキストボックス／図形／表／グラフ／ノートペイン）も記載する（合計50文字以内）
  ✅ 良い例: タイトル「サーバ構成」→「サーバ」
  ✅ 良い例: テキストボックス「処理は完了だ」→「だ」
  ❌ 悪い例: テキストボックス 2行目（どの語句が問題か不明）
- 指摘事項: 「何が問題か」を断言する一文で記述する
  ✅ 良い例: 「用語統一リスト（正: サーバー）と異なる表記が使用されている」
  ❌ 悪い例: 「表記ゆれが見受けられます」（曖昧）
- コードブロック（```）は出力に含めない

---

## スライド抽出データ

```json
{extract_json_str}
```

## 用語チェック結果

```json
{terminology_json_str}
```
"""


def _build_synthesis_prompt(
    filename: str,
    chunk_results: list[str],
    total_slides: int,
    reviewed_slides: list[int],
    selected_perspectives: list[dict] | None,
) -> str:
    """全チャンクの指摘結果を統合して最終レポートを生成するプロンプトを生成する。"""
    from datetime import date as _date
    today = _date.today().strftime("%Y-%m-%d")
    page_label = (
        "全スライド"
        if len(reviewed_slides) == total_slides
        else ", ".join(str(p) for p in reviewed_slides) + " ページ"
    )
    has_perspectives = bool(selected_perspectives)
    perspectives_count_row = _build_perspectives_count_row(selected_perspectives)

    all_findings = "\n\n---\n\n".join(
        f"【チャンク {i + 1}】\n{r}" for i, r in enumerate(chunk_results)
    )

    return f"""あなたはPowerPoint資料の品質レビュアーです。
以下は「{filename}」をチャンク分割してレビューした各チャンクの指摘結果です。
この結果をもとに、指定のMarkdownフォーマットで最終レポートを生成してください。

## 各チャンクの指摘結果

{all_findings}

---

## 出力フォーマット（厳守）

以下のMarkdownフォーマットを**厳密に**守ること。

# PowerPoint レビューレポート

**ファイル**: {filename}
**レビュー日時**: {today}
**総スライド数**: {total_slides} 枚（ファイル全体）
**レビュー対象ページ**: {page_label}
**対象読者**: 顧客向け

---

## カテゴリ別 指摘件数

| カテゴリ | 件数 |
|---------|------|
| 📝 表記ゆれ・文章校正 | N件 |
| 🔗 論理的整合性 | N件 |
{perspectives_count_row}
| 👥 文体 | N件 |
| **合計** | **N件** |

---

## スライド別 指摘事項

（各チャンクの指摘事項をスライド番号順に以下のテーブル形式で整理して列挙。指摘がないスライドは省略）

### スライド N：タイトル

| カテゴリ | 指摘事項 | 指摘箇所 |
|---------|---------|---------|
| 📝 | 用語統一リスト（正: サーバー）と異なる表記が使用されている | タイトル「サーバ構成」→「サーバ」 |

## 出力ルール
- 指摘箇所: 問題のあるテキストを**直接引用**し、場所（shape_kind: タイトル／テキストボックス／図形／表／グラフ／ノートペイン）も記載する（合計50文字以内）
  ✅ 良い例: タイトル「サーバ構成」→「サーバ」
  ❌ 悪い例: テキストボックス 2行目（どの語句が問題か不明）
- 指摘事項: 「何が問題か」を断言する一文で記述する
- カテゴリ別件数は各チャンクの指摘を正確に集計してから記載
- コードブロック（```）は出力に含めない。Markdownをそのまま出力する
"""


def _run_chunked_review(
    filename: str,
    extract_data: dict,
    terminology_data: dict,
    selected_perspectives: list[dict] | None,
    chunk_size: int,
) -> str:
    """チャンク分割による並列レビューを実行し、統合レポートを返す。

    chunk_size 以下のスライド数なら従来の一括レビューにフォールバックする。
    """
    slides = extract_data.get("slides", [])
    total_slides = extract_data.get("total_slides", len(slides))
    reviewed_slides = extract_data.get("reviewed_slides", [s["slide_number"] for s in slides])

    # チャンクサイズ以下なら一括レビュー
    if len(slides) <= chunk_size:
        logger.info("スライド数(%d)がチャンクサイズ(%d)以下のため一括レビュー", len(slides), chunk_size)
        return call_ai_review(
            build_review_prompt(filename, extract_data, terminology_data, selected_perspectives)
        )

    # チャンク分割
    chunks = _split_extract_data(extract_data, chunk_size)
    total_chunks = len(chunks)
    logger.info(
        "チャンク分割レビュー開始: %d枚 → %dチャンク（各%d枚）",
        len(slides), total_chunks, chunk_size,
    )

    # 並列にチャンクレビューを実行
    chunk_results: list[str] = [""] * total_chunks

    def _review_chunk(idx: int, chunk_data: dict) -> tuple[int, str]:
        logger.info("チャンク %d/%d レビュー中...", idx + 1, total_chunks)
        prompt = _build_chunk_prompt(
            filename, chunk_data, terminology_data,
            selected_perspectives, idx + 1, total_chunks,
        )
        result = call_ai_review(prompt)
        logger.info("チャンク %d/%d 完了", idx + 1, total_chunks)
        return idx, result

    with ThreadPoolExecutor(max_workers=min(total_chunks, 4)) as executor:
        futures = {executor.submit(_review_chunk, i, c): i for i, c in enumerate(chunks)}
        for future in as_completed(futures):
            idx, result = future.result()  # 401エラー等は RuntimeError として伝播
            chunk_results[idx] = result

    # 統合レポート生成
    logger.info("全チャンク完了。統合レポートを生成中...")
    synthesis_prompt = _build_synthesis_prompt(
        filename, chunk_results, total_slides, reviewed_slides, selected_perspectives
    )
    final_report = call_ai_review(synthesis_prompt)
    logger.info("統合レポート生成完了")
    return final_report


# ---------------------------------------------------------------------------
# ルート
# ---------------------------------------------------------------------------
@app.route("/")
def index():
    provider = detect_provider()
    provider_display = {"anthropic": "Claude", "openai": "OpenAI"}.get(provider, provider)
    return render_template("index.html", provider_display=provider_display)


@app.route("/debug")
def debug_info():
    """現在の設定状態・用語リストを確認できるデバッグエンドポイント。"""
    provider = detect_provider()

    # 用語リスト読み込み
    terminology_path = SKILL_DIR / "references" / "terminology.json"
    with _terminology_lock:
        try:
            with open(terminology_path, encoding="utf-8") as f:
                terminology = json.load(f)
            terms = terminology.get("terms", [])
        except Exception as e:
            terms = []
            terminology = {"error": str(e)}

    info = {
        "ai_provider": provider,
        "anthropic_api_key_set": bool(os.getenv("ANTHROPIC_API_KEY")),
        "openai_api_key_set": bool(os.getenv("OPENAI_API_KEY")),
        "openai_api_type": os.getenv("OPENAI_API_TYPE", "openai"),
        "openai_base_url": os.getenv("OPENAI_BASE_URL", ""),
        "openai_model": os.getenv("OPENAI_MODEL", "gpt-4o"),
        "azure_openai_endpoint": os.getenv("AZURE_OPENAI_ENDPOINT", ""),
        "azure_openai_api_version": os.getenv("AZURE_OPENAI_API_VERSION", "2024-02-01"),
        "scripts": {
            "extract_pptx": str(EXTRACT_SCRIPT),
            "extract_pptx_exists": EXTRACT_SCRIPT.exists(),
            "check_terminology": str(TERMINOLOGY_SCRIPT),
            "check_terminology_exists": TERMINOLOGY_SCRIPT.exists(),
        },
        "terminology": {
            "path": str(terminology_path),
            "exists": terminology_path.exists(),
            "version": terminology.get("version"),
            "last_updated": terminology.get("last_updated"),
            "term_count": len(terms),
            "terms": [
                {
                    "correct": t.get("correct"),
                    "variants": t.get("variants", []),
                    "category": t.get("category"),
                    "notes": t.get("notes"),
                }
                for t in terms
            ],
        },
    }
    return jsonify(info)


# terminology.json の並行読み書きを保護するロック
_terminology_lock = threading.Lock()

# ---------------------------------------------------------------------------
# 用語編集ロック（悲観的ロック）
# ---------------------------------------------------------------------------
import datetime

_EDIT_LOCK_TIMEOUT_SEC = 10 * 60  # 10分で自動解除

_edit_lock_state: dict = {
    "token": None,       # ロック保持者のトークン（None = 未ロック）
    "locked_at": None,   # ロック取得時刻（datetime）
}
_edit_lock_mutex = threading.Lock()  # _edit_lock_state 自体を保護するロック


def _edit_lock_is_expired() -> bool:
    """ロックが存在し、かつタイムアウト済みかどうかを返す（呼び出し前に _edit_lock_mutex を取得すること）。"""
    if _edit_lock_state["locked_at"] is None:
        return False
    elapsed = (datetime.datetime.now() - _edit_lock_state["locked_at"]).total_seconds()
    return elapsed >= _EDIT_LOCK_TIMEOUT_SEC


@app.route("/api/terminology/lock", methods=["GET"])
def get_edit_lock():
    """編集ロック状態を返す。"""
    with _edit_lock_mutex:
        if _edit_lock_state["token"] is None or _edit_lock_is_expired():
            # 期限切れなら自動解除
            _edit_lock_state["token"] = None
            _edit_lock_state["locked_at"] = None
            return jsonify({"locked": False})
        elapsed = (datetime.datetime.now() - _edit_lock_state["locked_at"]).total_seconds()
        remaining = max(0, int(_EDIT_LOCK_TIMEOUT_SEC - elapsed))
        return jsonify({"locked": True, "remaining_sec": remaining})


@app.route("/api/terminology/lock", methods=["POST"])
def acquire_edit_lock():
    """編集ロックを取得する。成功時はトークンを返す。"""
    with _edit_lock_mutex:
        # 期限切れロックは自動解除
        if _edit_lock_is_expired():
            _edit_lock_state["token"] = None
            _edit_lock_state["locked_at"] = None

        if _edit_lock_state["token"] is not None:
            elapsed = (datetime.datetime.now() - _edit_lock_state["locked_at"]).total_seconds()
            remaining = max(0, int(_EDIT_LOCK_TIMEOUT_SEC - elapsed))
            return jsonify({
                "ok": False,
                "message": f"現在他のユーザーが編集中です（あと約 {remaining // 60} 分 {remaining % 60} 秒で自動解除）",
            }), 409

        token = uuid.uuid4().hex
        _edit_lock_state["token"] = token
        _edit_lock_state["locked_at"] = datetime.datetime.now()
        logger.info("用語編集ロック取得 (token=%s...)", token[:8])
        return jsonify({"ok": True, "token": token})


@app.route("/api/terminology/lock", methods=["DELETE"])
def release_edit_lock():
    """編集ロックを解除する。token が一致する場合のみ解除。"""
    body = request.get_json(silent=True) or {}
    token = body.get("token", "")
    with _edit_lock_mutex:
        if _edit_lock_state["token"] is None:
            return jsonify({"ok": True, "message": "すでに解除済みです"})
        if _edit_lock_state["token"] != token:
            return jsonify({"ok": False, "message": "トークンが一致しません"}), 403
        _edit_lock_state["token"] = None
        _edit_lock_state["locked_at"] = None
        logger.info("用語編集ロック解除 (token=%s...)", token[:8])
        return jsonify({"ok": True})


def _load_terminology() -> dict:
    with _terminology_lock:
        try:
            with open(TERMINOLOGY_PATH, encoding="utf-8") as f:
                return json.load(f)
        except FileNotFoundError:
            return {"version": "1.0", "terms": []}


@app.route("/api/terminology", methods=["GET"])
def get_terminology():
    """用語リストを返す。"""
    data = _load_terminology()
    return jsonify(data)


@app.route("/api/terminology", methods=["POST"])
def save_terminology():
    """用語リストを保存する。"""
    body = request.get_json(silent=True)
    if body is None or "terms" not in body:
        return jsonify({"error": "terms フィールドが必要です。"}), 400

    terms = body["terms"]
    # バリデーション
    for i, t in enumerate(terms):
        if not isinstance(t.get("correct", ""), str) or not t.get("correct", "").strip():
            return jsonify({"error": f"terms[{i}].correct が空です。"}), 400
        if not isinstance(t.get("variants", []), list):
            return jsonify({"error": f"terms[{i}].variants はリストである必要があります。"}), 400

    with _terminology_lock:
        try:
            with open(TERMINOLOGY_PATH, encoding="utf-8") as f:
                existing = json.load(f)
        except FileNotFoundError:
            existing = {"version": "1.0", "terms": []}

        existing["terms"] = [
            {
                "correct": t["correct"].strip(),
                "variants": [v.strip() for v in t.get("variants", []) if str(v).strip()],
                "category": t.get("category", "").strip(),
                "notes": t.get("notes", "").strip(),
            }
            for t in terms
        ]
        existing["last_updated"] = date.today().isoformat()

        with open(TERMINOLOGY_PATH, "w", encoding="utf-8") as f:
            json.dump(existing, f, ensure_ascii=False, indent=2)

    logger.info("用語リスト保存: %d語", len(existing["terms"]))
    return jsonify({"ok": True, "term_count": len(existing["terms"])})


# ---------------------------------------------------------------------------
# 固有観点リスト API
# ---------------------------------------------------------------------------
_perspectives_lock = threading.Lock()

# 固有観点の編集ロック（用語ロックとは独立）
_persp_lock_state: dict = {
    "token": None,
    "locked_at": None,
}
_persp_lock_mutex = threading.Lock()


def _persp_lock_is_expired() -> bool:
    """固有観点ロックがタイムアウト済みかどうか。"""
    if _persp_lock_state["locked_at"] is None:
        return False
    elapsed = (datetime.datetime.now() - _persp_lock_state["locked_at"]).total_seconds()
    return elapsed >= _EDIT_LOCK_TIMEOUT_SEC


def _load_perspectives() -> dict:
    """固有観点リストを読み込む。"""
    with _perspectives_lock:
        try:
            with open(PERSPECTIVES_PATH, encoding="utf-8") as f:
                return json.load(f)
        except FileNotFoundError:
            return {"version": "1.0", "perspectives": []}


@app.route("/api/perspectives", methods=["GET"])
def get_perspectives():
    """固有観点リストを返す。"""
    data = _load_perspectives()
    return jsonify(data)


@app.route("/api/perspectives", methods=["POST"])
def save_perspectives():
    """固有観点リストを保存する。"""
    body = request.get_json(silent=True)
    if body is None or "perspectives" not in body:
        return jsonify({"error": "perspectives フィールドが必要です。"}), 400

    perspectives = body["perspectives"]
    # バリデーション
    for i, cat in enumerate(perspectives):
        if not isinstance(cat.get("category", ""), str) or not cat.get("category", "").strip():
            return jsonify({"error": f"perspectives[{i}].category が空です。"}), 400
        items = cat.get("items", [])
        if not isinstance(items, list):
            return jsonify({"error": f"perspectives[{i}].items はリストである必要があります。"}), 400
        for j, item in enumerate(items):
            if not isinstance(item.get("perspective", ""), str) or not item.get("perspective", "").strip():
                return jsonify({"error": f"perspectives[{i}].items[{j}].perspective が空です。"}), 400

    with _perspectives_lock:
        try:
            with open(PERSPECTIVES_PATH, encoding="utf-8") as f:
                existing = json.load(f)
        except FileNotFoundError:
            existing = {"version": "1.0", "perspectives": []}

        existing["perspectives"] = [
            {
                "category": cat["category"].strip(),
                "items": [
                    {
                        "perspective": item["perspective"].strip(),
                        "notes": item.get("notes", "").strip(),
                    }
                    for item in cat.get("items", [])
                    if item.get("perspective", "").strip()
                ],
            }
            for cat in perspectives
            if cat.get("category", "").strip()
        ]
        existing["last_updated"] = date.today().isoformat()

        with open(PERSPECTIVES_PATH, "w", encoding="utf-8") as f:
            json.dump(existing, f, ensure_ascii=False, indent=2)

    total_items = sum(len(c["items"]) for c in existing["perspectives"])
    logger.info("固有観点リスト保存: %dカテゴリ, %d観点", len(existing["perspectives"]), total_items)
    return jsonify({
        "ok": True,
        "category_count": len(existing["perspectives"]),
        "item_count": total_items,
    })


@app.route("/api/perspectives/lock", methods=["GET"])
def get_persp_lock():
    """固有観点の編集ロック状態を返す。"""
    with _persp_lock_mutex:
        if _persp_lock_state["token"] is None or _persp_lock_is_expired():
            _persp_lock_state["token"] = None
            _persp_lock_state["locked_at"] = None
            return jsonify({"locked": False})
        elapsed = (datetime.datetime.now() - _persp_lock_state["locked_at"]).total_seconds()
        remaining = max(0, int(_EDIT_LOCK_TIMEOUT_SEC - elapsed))
        return jsonify({"locked": True, "remaining_sec": remaining})


@app.route("/api/perspectives/lock", methods=["POST"])
def acquire_persp_lock():
    """固有観点の編集ロックを取得する。"""
    with _persp_lock_mutex:
        if _persp_lock_is_expired():
            _persp_lock_state["token"] = None
            _persp_lock_state["locked_at"] = None

        if _persp_lock_state["token"] is not None:
            elapsed = (datetime.datetime.now() - _persp_lock_state["locked_at"]).total_seconds()
            remaining = max(0, int(_EDIT_LOCK_TIMEOUT_SEC - elapsed))
            return jsonify({
                "ok": False,
                "message": f"現在他のユーザーが編集中です（あと約 {remaining // 60} 分 {remaining % 60} 秒で自動解除）",
            }), 409

        token = uuid.uuid4().hex
        _persp_lock_state["token"] = token
        _persp_lock_state["locked_at"] = datetime.datetime.now()
        logger.info("固有観点編集ロック取得 (token=%s...)", token[:8])
        return jsonify({"ok": True, "token": token})


@app.route("/api/perspectives/lock", methods=["DELETE"])
def release_persp_lock():
    """固有観点の編集ロックを解除する。"""
    body = request.get_json(silent=True) or {}
    token = body.get("token", "")
    with _persp_lock_mutex:
        if _persp_lock_state["token"] is None:
            return jsonify({"ok": True, "message": "すでに解除済みです"})
        if _persp_lock_state["token"] != token:
            return jsonify({"ok": False, "message": "トークンが一致しません"}), 403
        _persp_lock_state["token"] = None
        _persp_lock_state["locked_at"] = None
        logger.info("固有観点編集ロック解除 (token=%s...)", token[:8])
        return jsonify({"ok": True})


@app.route("/review", methods=["POST"])
def review():
    # --- バリデーション ---
    if "pptx_file" not in request.files:
        return jsonify({"error": "ファイルが選択されていません。"}), 400
    file = request.files["pptx_file"]
    if not file.filename or not file.filename.lower().endswith(".pptx"):
        return jsonify({"error": ".pptx ファイルのみ対応しています。"}), 400

    pages_raw = request.form.get("pages", "").strip()
    # チルダ範囲（1～5,10 や 3~5,11）をカンマ区切りに展開
    pages = expand_page_ranges(pages_raw) if pages_raw else None

    # --- 固有観点（JSON文字列で受信） ---
    selected_perspectives = None
    perspectives_raw = request.form.get("custom_perspectives", "").strip()
    if perspectives_raw:
        try:
            selected_perspectives = json.loads(perspectives_raw)
        except json.JSONDecodeError:
            logger.warning("固有観点のJSON解析に失敗: %s", perspectives_raw[:200])
            selected_perspectives = None

    # --- リクエスト固有ディレクトリを作成（スレッドセーフ） ---
    request_id = uuid.uuid4().hex
    request_dir = UPLOAD_DIR / request_id
    try:
        request_dir.mkdir(parents=True, exist_ok=True)

        with tempfile.NamedTemporaryFile(
            suffix=".pptx", dir=str(request_dir), delete=False
        ) as tmp:
            file.save(tmp.name)
            tmp_path = tmp.name
        logger.info("アップロード完了: %s (request_id=%s)", file.filename, request_id)

        # --- テキスト抽出 ---
        extract_data = run_extract(tmp_path, pages)

        # extract_pptx の stdout を文字列で再生成（check_terminology に渡すため）
        extract_stdout = json.dumps(extract_data, ensure_ascii=False, indent=2)

        # --- 用語チェック ---
        terminology_data = run_terminology_check(extract_stdout)

        # --- AI でレビュー（チャンク分割 + 並列実行 + 統合）---
        report_md = _run_chunked_review(
            file.filename, extract_data, terminology_data,
            selected_perspectives, _REVIEW_CHUNK_SIZE,
        )

        # --- 空レスポンス・異常コンテンツチェック（多重防御） ---
        if not report_md or not report_md.strip():
            logger.error("AIレビュー結果が空です。レポートの生成に失敗しました。")
            return jsonify({"error": "AIレビュー結果が空です。再度お試しください。"}), 500
        if len(report_md.strip()) < 50:
            logger.error(
                "AIレビュー結果が短すぎます（%d文字）。正常なレポートが生成されていない可能性があります。",
                len(report_md.strip()),
            )
            return jsonify({"error": "AIレビュー結果が不正です（内容が短すぎます）。再度お試しください。"}), 500
        if "# PowerPoint レビューレポート" not in report_md:
            logger.warning(
                "AIレビュー結果に期待するヘッダーが含まれていません。内容を確認してください（length=%d）",
                len(report_md),
            )

        # --- ダウンロード用ファイル名を生成（実ファイル保存なし・クライアント側でBlob生成）---
        stem = Path(file.filename).stem
        today_str = date.today().strftime("%Y%m%d")
        report_filename = f"{today_str}_{stem}_{request_id[:8]}.md"
        logger.info("レビュー完了: %s (request_id=%s)", file.filename, request_id)

        return jsonify(
            {
                "report": report_md,
                "report_filename": report_filename,
                "total_slides": extract_data.get("total_slides"),
                "reviewed_slides": extract_data.get("reviewed_slides", []),
            }
        )

    except RuntimeError as e:
        logger.error("レビューエラー: %s", e)
        return jsonify({"error": str(e)}), 500
    except Exception as e:
        logger.exception("予期しないエラー")
        return jsonify({"error": f"予期しないエラーが発生しました: {e}"}), 500
    finally:
        # ディレクトリごと削除（一時ファイルも含む）。ignore_errors=True でスレッドセーフ
        shutil.rmtree(request_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# OPENAI API キー設定エンドポイント
# ---------------------------------------------------------------------------
_env_file_lock = threading.Lock()


@app.route("/api/openai-key", methods=["POST"])
def set_openai_key():
    """OPENAI_API_KEY を .env ファイルに保存する。

    パスワード保護: config.yaml の api_key_setting.password と一致した場合のみ許可。
    パスワード未設定（空文字列）の場合は機能無効。
    """
    body = request.get_json(silent=True)
    if not body:
        return jsonify({"error": "リクエストボディが不正です。"}), 400

    password_input = body.get("password", "")
    api_key_input = body.get("api_key", "").strip()

    # パスワード検証
    cfg_password = _config.get("api_key_setting", {}).get("password", "").strip()
    if not cfg_password:
        return jsonify({"error": "API キー設定機能が無効です。config.yaml で password を設定してください。"}), 403
    if password_input != cfg_password:
        logger.warning("OPENAI KEY SETTING: パスワード不一致")
        return jsonify({"error": "パスワードが正しくありません。"}), 403

    # パスワードのみ検証（verify_only=true の場合はここで返す）
    if body.get("verify_only"):
        return jsonify({"ok": True, "message": "パスワードが確認されました。"})

    # APIキーのバリデーション
    if not api_key_input:
        return jsonify({"error": "API キーが空です。"}), 400

    # .env ファイルの OPENAI_API_KEY を更新（または追加）
    env_path = BASE_DIR / ".env"
    with _env_file_lock:
        if env_path.exists():
            lines = env_path.read_text(encoding="utf-8").splitlines(keepends=True)
        else:
            lines = []

        updated = False
        new_lines: list[str] = []
        for line in lines:
            # コメント行・空行・OPENAI_API_KEY 以外の行はそのまま保持
            stripped = line.strip()
            if stripped.startswith("OPENAI_API_KEY"):
                new_lines.append(f"OPENAI_API_KEY={api_key_input}\n")
                updated = True
            else:
                new_lines.append(line)

        if not updated:
            new_lines.append(f"OPENAI_API_KEY={api_key_input}\n")

        env_path.write_text("".join(new_lines), encoding="utf-8")

    # 実行中プロセスの環境変数も即時反映（再起動不要）
    os.environ["OPENAI_API_KEY"] = api_key_input
    logger.info("OPENAI_API_KEY を更新しました（キー末尾: ...%s）", api_key_input[-4:] if len(api_key_input) >= 4 else "****")
    return jsonify({"ok": True, "message": "OPENAI_API_KEY を更新しました。"})


def _sanitize_csv_value(value: str) -> str:
    """CSV injection 対策: Excel が数式として解釈しうるプレフィックスを無害化する。"""
    if value and value[0] in ("=", "+", "-", "@", "\t", "\r"):
        return "'" + value
    return value


_feedback_lock = threading.Lock()


@app.route("/api/feedback", methods=["POST"])
def api_feedback():
    """改善要望を受け取り、CSVファイルに追記する。"""
    body = request.get_json(silent=True) or {}
    content = (body.get("content") or "").strip()
    if not content:
        return jsonify({"error": "改善要望を入力してください。"}), 400

    name = (body.get("name") or "").strip()
    department = (body.get("department") or "").strip()

    if len(name) > 100:
        return jsonify({"error": "名前は100文字以内で入力してください。"}), 400
    if len(department) > 100:
        return jsonify({"error": "所属は100文字以内で入力してください。"}), 400
    if len(content) > 5000:
        return jsonify({"error": "改善要望は5000文字以内で入力してください。"}), 400

    now = datetime.datetime.now().strftime("%Y/%m/%d %H:%M")
    feedback_dir = BASE_DIR / "feedback"
    feedback_dir.mkdir(exist_ok=True)
    feedback_csv = feedback_dir / "feedback.csv"

    try:
        with _feedback_lock:
            write_header = not feedback_csv.exists()
            with feedback_csv.open("a", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                if write_header:
                    writer.writerow(["登録日時", "名前", "所属", "改善要望"])
                writer.writerow([
                    now,
                    _sanitize_csv_value(name),
                    _sanitize_csv_value(department),
                    _sanitize_csv_value(content),
                ])
    except OSError as e:
        logger.error("フィードバック CSV 書き込みエラー: %s", e)
        return jsonify({"error": "フィードバックの保存に失敗しました。"}), 500

    logger.info("改善要望を受け付けました: name=%s, dept=%s", name, department)
    return jsonify({"ok": True})


# ---------------------------------------------------------------------------
# エントリポイント
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="pptx-reviewer Web UI")
    parser.add_argument(
        "-p", "--port",
        type=int,
        default=None,
        help="ポート番号（デフォルト: 環境変数 PORT → 5000）",
    )
    args = parser.parse_args()

    # 優先順位: CLI引数 > 環境変数 PORT > デフォルト 5000
    port = args.port or int(os.getenv("PORT", 5000))
    debug = os.getenv("FLASK_DEBUG", "false").lower() == "true"

    # アクセスURLをわかりやすく表示
    try:
        local_ip = socket.gethostbyname(socket.gethostname())
    except Exception:
        local_ip = "x.x.x.x"

    logger.info("ローカル:    http://localhost:%d", port)
    logger.info("イントラネット: http://%s:%d", local_ip, port)

    try:
        from waitress import serve
        threads = int(os.getenv("WAITRESS_THREADS", 8))
        logger.info("waitress でサーバー起動 (port=%d, threads=%d)", port, threads)
        serve(app, host="0.0.0.0", port=port, threads=threads)
    except ImportError:
        logger.info("Flask 開発サーバーで起動 (port=%d, debug=%s)", port, debug)
        app.run(host="0.0.0.0", port=port, debug=debug)
