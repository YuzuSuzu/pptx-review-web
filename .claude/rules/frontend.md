---
paths:
  - "templates/*.html"
---

# Frontend Rules（index.html SPA）

## 基本制約

- `templates/index.html` は単一ファイルのSPA。分割しない
- CSSはすべて `<style>` タグ内に記述する。外部CSSファイルを追加しない
- JSはすべて `<script>` タグ内に記述する。外部JSファイルを追加しない（CDNライブラリ除く）

## カラーパレット（変更しない）

| 用途 | 値 |
|------|----|
| ヘッダー・メインカラー | `#1e3a5f` |
| アクセントブルー | `#3b82f6` |
| 成功グリーン | `#059669` |
| エラーレッド | `#b91c1c` |
| 背景 | `#f0f2f5` |
| カード背景 | `#fff` |

## モーダルのパターン

新しいモーダルを追加するときは既存のパターンに従う:

```html
<div class="modal-overlay" id="xxx-modal">
  <div class="modal-box" style="max-width:480px;">
    <div class="modal-title">アイコン タイトル</div>
    <div class="modal-body">
      <div id="xxx-alert" class="modal-alert"></div>
      <!-- フォーム要素 -->
    </div>
    <div class="modal-footer">
      <button class="btn-icon btn-cancel" id="xxx-modal-cancel">キャンセル</button>
      <button class="btn-icon btn-save"   id="xxx-modal-ok">OK</button>
    </div>
  </div>
</div>
```

JSはIIFEで包んで即時実行する:

```js
(function() {
  const modal = document.getElementById('xxx-modal');
  // 開く・閉じる・送信処理
})();
```

- オーバーレイ外クリックで閉じる（`e.target === modal` で判定）
- 成功時: OKボタンを非表示、キャンセルを「閉じる」に変更
- 通信中: ボタンを `disabled` にし、テキストを「送信中...」に変更

## レポート表示

- `marked.parse(markdownText)` でMarkdown→HTMLに変換して `#report-container` に挿入する
- テーブル・見出しのスタイルは `#report-container` 配下のCSSで定義済み。新たな上書きは原則しない

## ボタン種別

| クラス | 用途 |
|--------|------|
| `.btn-primary` | メインアクション（レビュー実行等） |
| `.btn-icon.btn-save` | モーダルOK・保存 |
| `.btn-icon.btn-cancel` | モーダルキャンセル |
| `.btn-icon.btn-edit` | 編集開始 |
| `.btn-feedback-open` | 改善要望（緑・`margin-left:auto` でヘッダー右寄せ） |
| `.btn-api-setting` | OPENAI KEY SETTING（グレー） |
