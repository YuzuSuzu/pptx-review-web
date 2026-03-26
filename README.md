# PPTX レビュアー Web UI

ブラウザから PowerPoint ファイル（.pptx）をアップロードするだけで、
AI が自動でレビューし、改善ポイントをレポートとして表示・保存するWebアプリです。

**Claude（Anthropic）と GPT-4o（OpenAI / Codex CLI 互換）の両方に対応しています。**

---

## このアプリで何ができるの？

> 「PowerPoint を作ったけど、提出前にチェックしてもらいたい」
> 「コマンドの操作は難しくて使えない」
> そんな方のために、**ブラウザだけで pptx-reviewer スキルを使える画面**を用意しました。

ブラウザでファイルを選んでボタンを押すだけで、以下の4点をAIが自動チェックします。

| チェック項目 | 内容 |
|------------|------|
| 📝 文章校正・表記ゆれ | 「サーバ」と「サーバー」が混在していないか、誤字脱字はないか |
| 🔗 論理的整合性 | 前のスライドと後のスライドで話が矛盾していないか |
| 📊 構成・可読性 | 1枚のスライドに情報が詰め込みすぎていないか |
| 👥 顧客向け表現 | 「RBAC」などの専門用語を説明なしに使っていないか |

レビュー結果はブラウザ上に表示され、Markdown ファイルとしてダウンロードもできます。

---

## 使用する AI の選び方

このWebアプリは、**どちらか一方の API キーがあれば動作します**。

| AI | 設定値 | 必要なもの | 特徴 |
|----|--------|-----------|------|
| Claude（Anthropic） | `AI_PROVIDER=anthropic` | Anthropic API キー | デフォルト設定 |
| GPT-4o（OpenAI） | `AI_PROVIDER=openai` | OpenAI API キー | Codex CLI と同じ AI エンジン |

> **Codex CLI との関係について**
> Codex CLI（OpenAI が提供するコマンドラインツール）は内部で OpenAI の API を使っています。
> そのため、OpenAI API キーがあれば、**Codex CLI をインストールしなくても**
> このWebアプリから同等のAIによるレビューが可能です。

---

## 仕組みの全体像

```
あなたのブラウザ
    │
    │  ① PPTXファイルをアップロード
    ▼
┌──────────────────────────────────────────────────────┐
│              Flask Webサーバー (app.py)               │
│                                                      │
│  ② extract_pptx.py                                   │  ← pptx-reviewer スキルのスクリプト
│     スライドのテキスト・構造情報を取り出す             │
│                                                      │
│  ③ check_terminology.py                              │  ← pptx-reviewer スキルのスクリプト
│     用語リスト（terminology.json）と照合して          │
│     表記ゆれを検出する                               │
│                                                      │
│  ④ AI によるレビュー分析（.envで切り替え）            │
│     ┌─────────────────────────────────┐              │
│     │ AI_PROVIDER=anthropic（デフォルト）│  Claude     │
│     │ AI_PROVIDER=openai               │  GPT-4o     │
│     └─────────────────────────────────┘              │
│     4つの観点でレビューレポートを生成                  │
│                                                      │
└──────────────────────────────────────────────────────┘
    │
    │  ⑤ Markdownレポートを表示・ダウンロード
    ▼
あなたのブラウザ
```

### 処理の流れをもう少し詳しく

1. **ファイルをアップロード** — ブラウザからPPTXを選択して「レビュー開始」を押す
2. **テキスト抽出** — `extract_pptx.py` がスライドの文字・構造情報をすべて取り出す
3. **用語チェック** — `check_terminology.py` が `references/terminology.json` と照合し、表記ゆれを検出する
4. **AIレビュー** — 抽出データをAIに送り、4つの観点で分析したレポートを生成する
5. **結果表示** — ブラウザにレポートが表示される。`web/uploads/` にMarkdownファイルも保存される

---

## セットとして使うスキル

このWebアプリは、**`pptx-reviewer` スキル**と一緒に使うことを前提に作られています。

```
C:\work\claude code_02\
├── skills\
│   └── pptx-reviewer\          ← このスキルのスクリプト・用語リストを使います
│       ├── scripts\
│       │   ├── extract_pptx.py         （テキスト抽出）
│       │   └── check_terminology.py    （用語チェック）
│       └── references\
│           └── terminology.json        （用語統一リスト）
│
└── web\                         ← ここがこのWebアプリ
    ├── app.py
    └── templates\
        └── index.html
```

> `pptx-reviewer` スキルが `skills/pptx-reviewer/` に配置されていないと動作しません。
> スキルの導入方法は [`skills/pptx-reviewer/README.md`](../skills/pptx-reviewer/README.md) をご覧ください。

---

## セットアップ手順

### 必要なもの

| 必要なもの | 確認方法 |
|-----------|---------|
| Python 3.8以上 | `python --version` |
| API キー（どちらか1つ） | 下記参照 |

---

### STEP 1: APIキーを取得する

使いたい AI に応じて、どちらか一方のキーを用意します。

#### Anthropic（Claude）を使う場合

1. https://console.anthropic.com/ にサインイン
2. 左メニュー「API Keys」→「Create Key」でキーを発行
3. `sk-ant-...` という形式のキーをコピーしておく

#### OpenAI を使う場合（Codex CLI 互換）

1. https://platform.openai.com/api-keys にサインイン
2. 「Create new secret key」でキーを発行
3. `sk-...` という形式のキーをコピーしておく

> OpenAI のキーがあれば、**Codex CLI をインストールしなくても**このWebアプリでGPT-4o によるレビューが可能です。

---

### STEP 2: APIキーを設定する

**`AI_PROVIDER` の設定は不要です。** APIキーを設定するだけで自動的にAIが選ばれます。

| 状況 | 自動選択されるAI |
|------|----------------|
| `OPENAI_API_KEY` のみ設定 | OpenAI（GPT-4o）|
| `ANTHROPIC_API_KEY` が設定されている | Claude（Anthropic）|
| 両方設定されている | Claude（Anthropic）優先 |

#### 方法A：環境変数で設定する（`.env` ファイル不要・お手軽）

Windowsのコマンドプロンプトでサーバー起動前に設定します。

```cmd
:: OpenAI を使う場合
set OPENAI_API_KEY=sk-xxxxxxxxxxxxxxxxxxxx

:: Anthropic を使う場合
set ANTHROPIC_API_KEY=sk-ant-xxxxxxxxxxxxxxxxxxxx
```

> `set` コマンドで設定した環境変数はそのターミナルを閉じるまで有効です。
> 毎回入力したくない場合は方法Bの `.env` ファイルを使ってください。

#### 方法B：`.env` ファイルで設定する（永続的に保存したい場合）

```cmd
cd "C:\work\claude code_02\web"
copy .env.example .env
```

作成した `.env` をメモ帳などで開き、使うキーだけ入力します。

```
# OpenAI を使う場合はこの行だけ入力
OPENAI_API_KEY=sk-xxxxxxxxxxxxxxxxxxxx

# Anthropic を使う場合はこの行だけ入力
ANTHROPIC_API_KEY=sk-ant-xxxxxxxxxxxxxxxxxxxx
```

モデルを変更したい場合のみ `OPENAI_MODEL` を追加してください（省略時は `gpt-4o`）。

```
OPENAI_MODEL=gpt-4o-mini
```

---

### STEP 3: 仮想環境を作ってライブラリをインストールする

仮想環境（venv）は「このアプリ専用のPython環境」です。
他のプロジェクトと干渉しないよう、必ずここで作ります。

```cmd
cd "C:\work\claude code_02\web"

python -m venv venv
venv\Scripts\pip install -r requirements.txt
```

> 「仮想環境って何？」という方へ：
> プロジェクトごとに必要なライブラリを入れる「専用の箱」だと思ってください。
> 一度作れば、次回からは `venv\Scripts\python app.py` と打つだけでOKです。

---

### STEP 4: サーバーを起動する

#### ポート番号の指定方法（3通り）

ポート番号は以下の優先順位で決まります。**上にあるほど優先されます。**

| 優先順位 | 方法 | 例 |
|---------|------|----|
| 1位（最優先） | 起動時の引数 `--port` / `-p` | `python app.py --port 8080` |
| 2位 | 環境変数 `PORT` | `set PORT=8080` |
| 3位（デフォルト） | 指定なし | ポート `5000` が使われる |

**引数で指定する場合（毎回変えたい時に便利）：**

```cmd
cd "C:\work\claude code_02\web"

:: デフォルト（ポート5000）
venv\Scripts\python app.py

:: ポートを指定して起動
venv\Scripts\python app.py --port 8080
venv\Scripts\python app.py -p 49156
```

**環境変数で指定する場合（毎回同じポートを使う時に便利）：**

```cmd
set PORT=8080
venv\Scripts\python app.py
```

**`.env` ファイルに書く場合（設定を永続化したい時）：**

```
PORT=8080
```

起動すると以下のようにアクセスURLが表示されます。

```
{"message": "ローカル:    http://localhost:8080"}
{"message": "イントラネット: http://192.168.1.11:8080"}
{"message": "waitress でサーバー起動 (port=8080)"}
```

---

### STEP 5: ブラウザで開く

#### ローカル（自分だけが使う場合）

```
http://localhost:5000
```

#### イントラネット（社内の他のPCからもアクセスしたい場合）

→ 次のセクション「**イントラネットでの公開方法**」を参照してください。

---

## イントラネットでの公開方法

このアプリは最初から `0.0.0.0`（すべてのネットワークインターフェース）でリッスンしているため、
**ポートを開放するだけで社内の他のPCからアクセスできます。**

### よく使われるポート番号の例

ポートを選ぶ際の参考にしてください。

| ポート番号 | 区分 | 用途・特徴 |
|-----------|------|-----------|
| `5000` | ダイナミック | **このアプリのデフォルト**。Flask の慣例的なポート |
| `8000` | ダイナミック | Python の `http.server` や Django 開発サーバーでよく使われる |
| `8080` | ダイナミック | HTTP の代替として広く使われる。社内アプリに多い |
| `8888` | ダイナミック | Jupyter Notebook のデフォルト |
| `3000` | ダイナミック | Node.js / React 開発サーバーでよく使われる |
| `49152`〜`65535` | ダイナミック/プライベート | OS が自由に使えるポート範囲。社内用途で衝突が少ない |
| `49156` | ダイナミック/プライベート | 上記範囲内の任意ポート。他アプリと被りにくい |
| `80` | ウェルノウン | HTTP 標準。使用には管理者権限が必要な場合がある |
| `443` | ウェルノウン | HTTPS 標準。証明書の設定が必要 |

> **社内利用のおすすめ**: `8080` または `49152〜65535` の範囲から選ぶと他のソフトウェアと衝突しにくいです。
> 社内にすでに別のサービスが動いている場合は IT 管理者に空きポートを確認してください。

---

### STEP 1: このPCのIPアドレスを確認する

コマンドプロンプトで以下を実行します。

```cmd
ipconfig
```

`IPv4 アドレス` の行に表示される値がこのPCのIPアドレスです。

```
イーサネット アダプター:
   IPv4 アドレス . . . . . . . . : 10.100.100.100   ← これ
```

起動時にもイントラネットURLが自動表示されます（`ipconfig` を省略できます）。

```
{"message": "イントラネット: http://10.100.100.100:8080"}
```

### STEP 2: ポートを指定してサーバーを起動する

```cmd
cd "C:\work\claude code_02\web"

venv\Scripts\python app.py --port 8080
```

### STEP 3: Windowsファイアウォールでポートを開放する

社内の他のPCからアクセスするには、Windowsファイアウォールで該当ポートへの接続を許可する必要があります。

**管理者権限のコマンドプロンプト**で以下を実行します（ポート番号は実際の値に変えてください）。

```cmd
netsh advfirewall firewall add rule name="pptx-reviewer" protocol=TCP dir=in localport=8080 action=allow
```

> ファイアウォールの設定変更はIT管理者に確認してから行ってください。

### STEP 4: 他のPCからアクセスする

サーバーPCの IPアドレスとポートを使って、他のPCのブラウザからアクセスします。

```
http://10.100.100.100:8080
```

| 項目 | 例 |
|------|-----|
| サーバーPCのIP | `10.100.100.100`（起動ログまたは `ipconfig` で確認） |
| ポート | `8080`（`--port` で指定した値） |
| アクセスURL | `http://10.100.100.100:8080` |

---

## 使い方

1. **ファイルを選択** — 「クリックまたはドラッグ」エリアにPPTXをドロップするか、クリックして選択する
2. **ページを指定（任意）** — 「対象ページ」に `1,3,5` のように入力すると、指定したページだけをレビュー
3. **「レビュー開始」を押す** — AIが分析中はプログレスメッセージが表示される（1〜2分程度）
4. **レポートを確認** — 画面にレポートが表示される
5. **保存** — 「Markdown を保存」ボタンでファイルをダウンロード。`web/uploads/` にも自動保存される

---

## フォルダ構成

```
web\
├── app.py               ← Flaskサーバー本体
├── requirements.txt     ← 必要なPythonライブラリの一覧
├── .env.example         ← APIキー設定のテンプレート
├── .env                 ← 実際のAPIキー（自分で作成・gitには含めない）
├── .gitignore
├── venv\                ← 仮想環境（自動生成）
├── uploads\             ← アップロードした一時ファイルとレポートの保存先
└── templates\
    └── index.html       ← ブラウザに表示されるUI
```

---

## よくある質問

**Q. サーバーを止めるには？**
ターミナルで `Ctrl + C` を押すと停止できます。

**Q. イントラネットで公開できますか？**
できます。`python app.py` で起動した時点で `0.0.0.0` でリッスンしているため、
Windowsファイアウォールでポートを開放すれば社内の他のPCから `http://<このPCのIP>:<ポート>` でアクセスできます。
詳しくは「イントラネットでの公開方法」セクションを参照してください。

**Q. Streamlit でも同じことができますか？**
Streamlit でも `streamlit run app.py --server.address 0.0.0.0` でイントラネット公開できます。
このアプリは Flask + waitress で同等の構成を実現しており、追加ツールは不要です。

**Q. レポートはどこに保存される？**
`web/uploads/` フォルダに `YYYYMMDD_ファイル名.md` という名前で保存されます。

**Q. Codex CLI がなくてもOpenAIのAIを使えますか？**
はい。OpenAI API キーがあれば、Codex CLI をインストールしなくても
`.env` に `AI_PROVIDER=openai` と `OPENAI_API_KEY` を設定するだけで使えます。
Codex CLI も内部ではOpenAI APIを使っているため、レビュー結果は同等です。

**Q. AnthropicとOpenAI、どちらのAIを使えばいい？**
どちらでも問題なくレビューできます。すでにどちらかのAPIキーを持っていれば、そちらを使うのが簡単です。

**Q. 大きいファイルはアップロードできますか？**
最大 **50MB** まで対応しています。

**Q. 用語リストをカスタマイズしたい**
`skills/pptx-reviewer/references/terminology.json` を編集してください。
書き方は `pptx-reviewer` の README をご参照ください。
