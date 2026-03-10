# トラブルシューティング

## `refresh.js` WebSocket `ws://localhost:8081/` 接続エラー

Live Server や Cursor のプレビュー等が、自動リロード用に `refresh.js` を注入し、WebSocket で `localhost:8081` に接続します。**8081 でリロードサーバーを起動していない**場合、コンソールに接続失敗が出ます。

**対処:**

- ローカルで HTML だけ開いて動作確認する場合: 拡張機能の Live Preview / Live Server をオフにするか、`file://` または静的サーバーで **注入なし** で開く。
- 開発時にリロードが不要なら: プレビューを閉じるか、別ブラウザで GAS デプロイ URL や GitHub Pages だけを開く。
- エラー自体はアプリの GAS 連携とは無関係で、**無視して問題ない**ことも多いです。

## 打刻時「データの行数が範囲の行数と一致しません」

スタッフ別シートへ `setValues` するとき、`getRange` の第3・第4引数は **終了行・終了列ではなく行数・列数** です。誤って `getRange(nextRow, 1, nextRow, 4)` と書くと、行数が `nextRow`（例: 1002）扱いになり 1 行分のデータと不一致になります。

**修正済み:** `Code.js` では `getRange(nextRow, 1, 1, 4)` で 1 行×4 列に限定しています。
