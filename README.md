# TSUMUGI Order App

紙の「TSUMUGI ORDER CONCEPT SHEET」（3ページ）を、1画面1問のウィザード形式に再構成したセルフ記入アプリです。お客様がスマホで開き、店舗コンセプト・器の仕様・製作スケジュールを順に入力すると、[Google Apps Script](apps-script/) 経由でスプレッドシート・Drive・メールに送信されます。

サーバー・ビルドステップなし。静的な `index.html` + `support.js`（ランタイム）+ デザインシステムの CSS/JS だけで動作し、GitHub Pages でそのまま公開できます。

## 構成

- `index.html` — アプリ本体（ウィザードUI + ロジック）
- `support.js` — テンプレート/状態管理ランタイム（CDN から React・ReactDOM を読み込む）
- `_ds/modernist-.../` — デザインシステムのトークンとコンポーネントCSS
- `apps-script/` — 送信先となる Google Apps Script のコードと設置手順（[SETUP.md](apps-script/SETUP.md)）

## ローカルで確認する

ビルド不要。任意の静的サーバーで配信するだけです。

```bash
python3 -m http.server 8000
```

`http://localhost:8000` を開くとアプリが起動します（送信先未設定時はデモ送信になります）。

## デプロイ（GitHub Pages）

このリポジトリの Pages を有効にすると、`main` ブランチの `index.html` がそのまま公開されます。公開後、実際に送信データを受け取るには [apps-script/SETUP.md](apps-script/SETUP.md) の手順で Google Apps Script を設置し、発行された Web App URL をアプリの Tweaks（`endpoint` プロパティ）に設定してください。
