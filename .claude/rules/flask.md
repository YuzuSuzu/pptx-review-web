---
paths:
  - "app.py"
  - "tests/**/*.py"
---

# Flask Rules

## エンドポイント追加時のパターン

- レスポンスは必ず `jsonify()` を使う。辞書を直接返さない
- 成功時は `{"ok": True, ...}`、エラー時は `{"error": "メッセージ"}` + HTTPステータスコード
- `request.get_json(silent=True) or {}` で body を受け取る（ParseError を防ぐ）
- バリデーション失敗は 400、認証失敗は 403 を返す

```python
@app.route("/api/xxx", methods=["POST"])
def api_xxx():
    body = request.get_json(silent=True) or {}
    value = (body.get("key") or "").strip()
    if not value:
        return jsonify({"error": "..."}), 400
    # 処理
    return jsonify({"ok": True})
```

## ファイルアップロード処理

- アップロードファイルは必ず `uuid.uuid4()` のサブディレクトリに保存し、`finally` ブロックで削除する
- ファイルサイズ上限は `app.config["MAX_CONTENT_LENGTH"]`（現在 50 MB）で制御

## スキルスクリプトの呼び出し

- `run_extract()` / `run_terminology_check()` を使い、subprocess の直呼び出しは避ける
- これらの関数は `SKILL_DIR` を自動解決してスクリプトを実行する

## 排他制御

- 用語リスト編集: `GET /api/terminology/lock`（状態確認）→ `POST /api/terminology/lock`（取得）→ `POST /api/terminology`（保存）→ `DELETE /api/terminology/lock`（解放）
- 固有観点編集: `GET /api/perspectives/lock` → `POST /api/perspectives/lock` → `POST /api/perspectives` → `DELETE /api/perspectives/lock`（同パターン）
- `.env` ファイルへの書き込みは `_env_file_lock`（`threading.Lock`）を使う

## テスト

- AIコール（`call_ai_review` / `call_claude_review` / `call_openai_review`）は `unittest.mock.patch` でモックする
- テスト用PPTXは `../skills/pptx-reviewer/test-files/dummy_proposal.pptx` を優先的に使用する
