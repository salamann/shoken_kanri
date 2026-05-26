# shoken_kanri

Google Apps Script ベースの資産管理/保有数管理プロジェクトです。

## 概要

- `email_notification.js` でスプレッドシートの週報を HTML メール送信します。
- `update_shoken.js` で保有株数の更新フォームを管理します。
- `holdingsForm.html` は保有数更新用のフォーム UI です。
- `index.html` はプロジェクトの HTML テンプレートとして利用されます。

## 主な機能

- `sendSheetByMail()`:
  - `メール用` シートやチャートを読み込み、HTML メールを生成して送信
  - `GEMINI_API_KEY` を使って AI コメントを生成
- `getCurrentHoldings()` / `updateHoldings()`:
  - `銘柄一覧` シートから保有情報を読み込み
  - フォーム入力で保有数を更新

## ファイル構成

- `appsscript.json` - GAS プロジェクト設定
- `email_notification.js` - 週報メール生成と送信ロジック
- `update_shoken.js` - 保有数管理ロジック
- `holdingsForm.html` - 保有数更新フォーム UI
- `index.html` - HTML テンプレート
- `email_address.js` - 送信先メールアドレス
- `.gitignore` - `email_address.js` と `node_modules/` を除外

## セットアップ

1. Google Apps Script プロジェクトを作成し、ファイルをアップロード / コピペ
2. `email_address.js` を以下のように作成する

```js
const EMAIL_TO = "your_email@example.com";
```

3. `email_address.js` は `.gitignore` に追加済みのため、Git に含めないようになります。
4. `GEMINI_API_KEY` は GAS の `PropertiesService.getScriptProperties()` に設定してください。

## 使い方

- GAS 側から `sendSheetByMail()` を実行して週報メール送信
- `holdingsForm.html` を表示して保有数を更新

## 注意事項

- `email_address.js` などの公開したくない情報は Git 管理しないでください。
- 実際の保有データは Google スプレッドシート内の `銘柄一覧` シートに保存されるため、スプレッドシートの共有設定にも注意が必要です。
- `appsscript.json` の `webapp.access` が `ANYONE_ANONYMOUS` になっている場合は、Web アプリ化時の公開範囲を確認してください。
