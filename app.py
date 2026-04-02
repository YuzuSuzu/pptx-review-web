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
from datetime import date
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
    known_terms: list[str] | None = None,
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

    # 顧客向け表現の調整観点：custom_perspectives.json 登録用語は既知のため指摘しない
    known_terms_instruction = ""
    if known_terms:
        terms_str = "、".join(known_terms)
        known_terms_instruction = (
            f"- 以下の用語は顧客が既に知っているため指摘しない：{terms_str}\n"
        )

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

### 観点2：論理的整合性
- スライド間の矛盾（前後で事実が食い違っていないか）
- 主語・述語のねじれ
- 因果関係の破綻（接続詞と内容が対応していないケース）

{perspectives_section}### 観点{perspective_count}：顧客向け表現の調整
{known_terms_instruction}- 技術的略語（APIM、RBAC、IaC 等）を説明なしに使用している箇所を指摘（ただし上記の既知用語は除く）
- です・ます調とだ・である調の混在
- 「など」「等」「場合によっては」の多用

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

## 総合サマリー

（全体を通じた主な課題と優先度が高い改善ポイントを3〜5行で要約。表記ゆれの件数・論理問題の有無・{"固有観点の問題・" if has_perspectives else ""}専門用語の多用・文体の乱れなど）

---

## カテゴリ別 指摘件数

| カテゴリ | 件数 |
|---------|------|
| 📝 表記ゆれ・文章校正 | N件 |
| 🔗 論理的整合性 | N件 |
{perspectives_count_row}
| 👥 専門用語・顧客表現 | N件 |
| **合計** | **N件** |

---

## スライド別 指摘事項

### スライド N：タイトル

| カテゴリ | 箇所 | 指摘内容 | 改善案 |
|---------|------|---------|--------|
| 📝 表記ゆれ | テキストボックス 2行目 | ... | ...することを推奨 |

（指摘がないスライドは省略）

---

## 全体的な改善提案

1. 改善提案1
2. 改善提案2
3. 改善提案3
```

## 出力ルール
- テキストの引用・転載は禁止。指摘内容と箇所の説明のみ記載
- 箇所の示し方: shape_kind（タイトル／テキストボックス／図形／表／グラフ／ノートペイン）と位置（「○行目」など）
- 改善案は「〜することを推奨」で統一
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
    """API 認証エラー（401）発生時に Teams Incoming Webhook で通知する。"""
    teams_cfg = _config.get("teams", {})
    if not teams_cfg.get("enabled"):
        return
    webhook_url = teams_cfg.get("webhook_url", "").strip()
    if not webhook_url:
        logger.warning("Teams 通知: webhook_url が config.yaml に設定されていません")
        return

    import datetime as _dt

    payload = {
        "@type": "MessageCard",
        "@context": "http://schema.org/extensions",
        "themeColor": "FF0000",
        "summary": "API 認証エラー (401)",
        "sections": [
            {
                "activityTitle": "⚠️ pptx-reviewer: API 接続エラー (Error code: 401)",
                "activitySubtitle": "APIキーの認証に失敗しました。.env ファイルのAPIキーを確認してください。",
                "facts": [
                    {"name": "ホスト", "value": socket.gethostname()},
                    {"name": "発生時刻", "value": _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
                    {"name": "エラー詳細", "value": error_message[:500]},
                ],
            }
        ],
    }
    try:
        resp = requests_lib.post(webhook_url, json=payload, timeout=10)
        resp.raise_for_status()
        logger.info("Teams 通知送信完了")
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
        raise RuntimeError(f"API認証エラー (Error code: 401): {err_str}") from e
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
            raise RuntimeError(f"API認証エラー (Error code: 401): {err_str}") from e
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


def call_ai_review(prompt: str) -> str:
    """検出したプロバイダーに応じて Anthropic または OpenAI を呼び出す。空レスポンス時は1回リトライ。"""
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
# ルート
# ---------------------------------------------------------------------------
@app.route("/")
def index():
    return render_template("index.html")


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

    pages = request.form.get("pages", "").strip() or None

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

        # --- AI でレビュー（AI_PROVIDER に応じて Anthropic / OpenAI を切り替え）---
        # custom_perspectives.json の全登録用語を「顧客既知用語」として収集
        all_persp_data = _load_perspectives()
        known_terms: list[str] = [
            item["perspective"].strip()
            for cat in all_persp_data.get("perspectives", [])
            for item in cat.get("items", [])
            if item.get("perspective", "").strip()
        ]
        prompt = build_review_prompt(
            file.filename, extract_data, terminology_data, selected_perspectives,
            known_terms=known_terms if known_terms else None,
        )
        report_md = call_ai_review(prompt)

        # --- 空レスポンスチェック（二重防御） ---
        if not report_md or not report_md.strip():
            logger.error("AIレビュー結果が空です。レポートの生成に失敗しました。")
            return jsonify({"error": "AIレビュー結果が空です。再度お試しください。"}), 500

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
