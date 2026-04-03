---
paths:
  - "app.py"
---

# AIレビュープロンプト・チャンク処理 Rules

## プロンプト構造（変更時の注意）

`build_review_prompt()` と `_build_synthesis_prompt()` は出力フォーマットを厳密に指定している。
変更する場合は `_TEST_MODE_RESPONSE` 定数（モックレスポンス）も同じフォーマットに合わせること。

### カテゴリと絵文字の対応

| 絵文字 | カテゴリ名 | 条件 |
|--------|-----------|------|
| 📝 | 文章校正・表記ゆれ（観点1） | 常時 |
| 🔗 | 論理的整合性（観点2） | 常時 |
| 🎯 | 固有観点（観点3） | 固有観点選択時のみ |
| 👥 | 文体（観点N） | 常時 |

- **カテゴリ別件数テーブル**: 絵文字＋カテゴリ名を記載する（例: `📝 文章校正・表記ゆれ`）
- **スライド別指摘事項テーブル**: カテゴリ列は絵文字のみ
- 注意: `_TEST_MODE_RESPONSE` は現在 📊（情報量・視認性）を使用しているが、これはプロンプト定義外の旧カテゴリ。カテゴリ変更時は `_TEST_MODE_RESPONSE` も合わせて更新すること

### 指摘禁止事項（プロンプトに明記済み）

箇条書き・体言止め・短文はスライドの標準書式であり指摘対象外。
「文が断片的」「見出し的項目が連続」などの指摘を追加しない。

## チャンク処理

`_run_chunked_review()` が主制御。以下の閾値を変えたい場合は `config.yaml` か `REVIEW_CHUNK_SIZE` 環境変数で対応し、コードを直接変更しない:

```yaml
review:
  chunk_size: 3  # 1リクエストあたりのスライド数
```

- スライド数 ≦ chunk_size → `build_review_prompt()` で一括レビュー
- スライド数 > chunk_size → `_build_chunk_prompt()` で並列チャンク + `_build_synthesis_prompt()` で統合

並列実行は `ThreadPoolExecutor(max_workers=min(チャンク数, 4))` で動的に決定。チャンク結果はスライド番号順に結合してから統合プロンプトへ渡す。

## TEST_MODE

`TEST_MODE=true` のとき `call_ai_review()` はAPIを呼ばず `_TEST_MODE_RESPONSE` を返す。
プロンプト変更時はこのモックレスポンスもあわせて更新する。

## 401エラー時の挙動

`call_claude_review()` / `call_openai_review()` が 401 を受け取った場合:
1. Teams webhook へ通知（`config.yaml` の `teams.enabled: true` の場合）
2. ユーザー向けに「APIキーの認証に失敗しました。管理者に連絡してください。」を返す
