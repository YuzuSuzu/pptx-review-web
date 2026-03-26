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
import os
import subprocess
import sys
import tempfile
from datetime import date
from pathlib import Path

import anthropic
from dotenv import load_dotenv
from flask import Flask, jsonify, render_template, request

try:
    from openai import OpenAI as OpenAIClient
    _OPENAI_AVAILABLE = True
except ImportError:
    _OPENAI_AVAILABLE = False

# ---------------------------------------------------------------------------
# ログ設定
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format='{"timestamp": "%(asctime)s", "level": "%(levelname)s", "message": "%(message)s"}',
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# 設定
# ---------------------------------------------------------------------------
load_dotenv()

BASE_DIR = Path(__file__).parent


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
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

MAX_UPLOAD_MB = 50

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024


# ---------------------------------------------------------------------------
# レビュープロンプト生成
# ---------------------------------------------------------------------------
def build_review_prompt(filename: str, extract_data: dict, terminology_data: dict) -> str:
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

    return f"""あなたはPowerPoint資料の品質レビュアーです。
以下のスライド抽出データと用語チェック結果をもとに、4つの観点でレビューを行い、
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

### 観点3：構成と可読性
- total_chars が 300文字超はやや多め、600文字超は要注意として指摘
- 見出し階層（level）が飛んでいないか
- 1スライド1メッセージ原則（複数テーマ混在はスライド分割を提案）

### 観点4：顧客向け表現の調整
- 技術的略語（APIM、RBAC、IaC 等）を説明なしに使用している箇所を指摘
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

（全体を通じた主な課題と優先度が高い改善ポイントを3〜5行で要約。表記ゆれの件数・論理問題の有無・専門用語の多用・文体の乱れなど）

---

## カテゴリ別 指摘件数

| カテゴリ | 件数 |
|---------|------|
| 📝 表記ゆれ・文章校正 | N件 |
| 🔗 論理的整合性 | N件 |
| 📊 構成・可読性 | N件 |
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
    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=8192,
        messages=[{"role": "user", "content": prompt}],
    )
    return message.content[0].text


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

    response = client.chat.completions.create(
        model=model,
        max_tokens=8192,
        messages=[{"role": "user", "content": prompt}],
    )
    return response.choices[0].message.content


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
    """検出したプロバイダーに応じて Anthropic または OpenAI を呼び出す。"""
    provider = detect_provider()
    logger.info("AI プロバイダー: %s", provider)
    if provider == "openai":
        return call_openai_review(prompt)
    return call_claude_review(prompt)


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


TERMINOLOGY_PATH = SKILL_DIR / "references" / "terminology.json"


def _load_terminology() -> dict:
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

    existing = _load_terminology()
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


@app.route("/review", methods=["POST"])
def review():
    # --- バリデーション ---
    if "pptx_file" not in request.files:
        return jsonify({"error": "ファイルが選択されていません。"}), 400
    file = request.files["pptx_file"]
    if not file.filename or not file.filename.lower().endswith(".pptx"):
        return jsonify({"error": ".pptx ファイルのみ対応しています。"}), 400

    pages = request.form.get("pages", "").strip() or None

    # --- 一時ファイルに保存 ---
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            suffix=".pptx", dir=str(UPLOAD_DIR), delete=False
        ) as tmp:
            file.save(tmp.name)
            tmp_path = tmp.name
        logger.info("アップロード完了: %s -> %s", file.filename, tmp_path)

        # --- テキスト抽出 ---
        extract_data = run_extract(tmp_path, pages)

        # extract_pptx の stdout を文字列で再生成（check_terminology に渡すため）
        extract_stdout = json.dumps(extract_data, ensure_ascii=False, indent=2)

        # --- 用語チェック ---
        terminology_data = run_terminology_check(extract_stdout)

        # --- AI でレビュー（AI_PROVIDER に応じて Anthropic / OpenAI を切り替え）---
        prompt = build_review_prompt(file.filename, extract_data, terminology_data)
        report_md = call_ai_review(prompt)

        # --- レポートをPPTXと同じ場所に保存（uploads/直下） ---
        stem = Path(file.filename).stem
        today_str = date.today().strftime("%Y%m%d")
        report_filename = f"{today_str}_{stem}.md"
        report_path = UPLOAD_DIR / report_filename
        if report_path.exists():
            i = 1
            while True:
                candidate = UPLOAD_DIR / f"{today_str}_{stem}_{i:02d}.md"
                if not candidate.exists():
                    report_path = candidate
                    report_filename = candidate.name
                    break
                i += 1
        report_path.write_text(report_md, encoding="utf-8")
        logger.info("レポート保存: %s", report_path)

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
        if tmp_path and Path(tmp_path).exists():
            Path(tmp_path).unlink()


# ---------------------------------------------------------------------------
# エントリポイント
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import argparse
    import socket

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
        logger.info("waitress でサーバー起動 (port=%d)", port)
        serve(app, host="0.0.0.0", port=port)
    except ImportError:
        logger.info("Flask 開発サーバーで起動 (port=%d, debug=%s)", port, debug)
        app.run(host="0.0.0.0", port=port, debug=debug)
