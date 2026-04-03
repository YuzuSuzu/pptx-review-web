# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## コマンド

```bash
# 初回セットアップ
python -m venv venv
venv/Scripts/pip install -r requirements.txt
cp .env.example .env  # APIキーを記入する

# サーバー起動
venv/Scripts/python app.py                    # port 5000
venv/Scripts/python app.py --port 8080
TEST_MODE=true venv/Scripts/python app.py     # APIキー不要・モック応答

# テスト
mkdir -p test-results
venv/Scripts/pytest tests/ -v 2>&1 | tee test-results/result_$(date +%Y%m%d_%H%M%S).log
venv/Scripts/pytest tests/test_app.py::test_review_with_mock -v  # 単一テスト
```

## アーキテクチャ

### ファイル構成
- `app.py` — Flask バックエンド（全ロジック集約）
- `templates/index.html` — SPA フロントエンド（marked.js でMarkdown→HTMLレンダリング）
- `config.yaml` — ログ保持日数・チャンクサイズ・Teams通知・APIキー設定パスワード
- `.env` — APIキー・ポート等（`.env.example` が雛形）
- `tests/test_app.py` — Flaskテストクライアント + AIモック使用のE2Eテスト

### スキルディレクトリ（外部依存）
`pptx-reviewer` スキルのスクリプトを subprocess で呼び出す。`SKILL_DIR` の解決順:
1. 環境変数 `SKILL_DIR`
2. `~/.codex/skills/pptx-review-text`（Codex CLI）
3. `~/.claude/skills/pptx-reviewer`（Claude Code）
4. `../skills/pptx-reviewer`（開発用相対パス）

使用するスキルファイル: `scripts/extract_pptx.py`、`scripts/check_terminology.py`、`references/terminology.json`、`references/custom_perspectives.json`

### AIプロバイダー自動検出
`detect_provider()` の優先順位:
1. 環境変数 `AI_PROVIDER` が明示指定されていればそれを使用
2. `OPENAI_API_KEY` のみ存在 → `openai`
3. `ANTHROPIC_API_KEY` が存在 → `anthropic`（両方ある場合も anthropic 優先）

OpenAI 系の接続先は `OPENAI_API_TYPE` で切り替え:
- `openai`（デフォルト）: 標準 OpenAI API。`OPENAI_BASE_URL` を設定すると LiteLLM 等の互換プロキシも使用可能
- `azure`: Azure OpenAI Service。`AZURE_OPENAI_ENDPOINT` が必須

### 主要な設定値

| 設定 | 場所 | 備考 |
|------|------|------|
| `AI_PROVIDER` | `.env` | 環境変数 > 自動検出（`anthropic` / `openai`） |
| `OPENAI_API_TYPE` | `.env` | `openai`（デフォルト）/ `azure` |
| `OPENAI_BASE_URL` | `.env` | LiteLLM等互換プロキシURL（openai型のみ） |
| `AZURE_OPENAI_ENDPOINT` | `.env` | azure型のとき必須 |
| `AZURE_OPENAI_API_VERSION` | `.env` | デフォルト: `2024-02-01` |
| `REVIEW_CHUNK_SIZE` | `.env` / `config.yaml` | 環境変数 > yaml（デフォルト: 3） |
| `LOG_RETENTION_DAYS` | `.env` / `config.yaml` | 環境変数 > yaml（デフォルト: 7日） |
| `PORT` | `--port` 引数 / `.env` | 引数 > 環境変数 > 5000 |
| APIキー設定パスワード | `config.yaml` `api_key_setting.password` | 空文字=機能無効 |
| Teams通知 | `config.yaml` `teams` | `power_automate` または `incoming_webhook` |
