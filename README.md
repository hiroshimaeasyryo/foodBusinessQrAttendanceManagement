# AccessControlFoodBusiness（出勤記録）

飲食系スタッフの出勤管理用アプリ。フォーク元の入退室管理ではなく、**出勤時刻・退勤時刻・休憩時間**を定時で記録する仕様です。

## 仕様概要

- **認証**: 誕生日4桁（MMdd）入力 → 候補スタッフ表示 → スタッフ選択 → QRスキャン
- **記録**: 属性（社員 / 契約社員 / 業務委託）に応じたデフォルト時刻を記録。イレギュラー時は任意入力。
- **スプレッドシート**: 指定IDのスプレッドシートに「打刻記録」および各スタッフ名のシートへ書き込み。

詳細は [docs/implementation.md](docs/implementation.md) を参照してください。

## 構成

| ファイル | 説明 |
|----------|------|
| `Code.js` | GAS バックエンド（スタッフ取得・打刻記録・キャッシュ） |
| `index.html` | GAS で配信する画面（テンプレート） |
| `standalone.html` | GitHub Pages 等でホスティングする単体HTML |
| `.clasp.json` | clasp 用設定（要 scriptId 設定） |
| `appsscript.json` | GAS プロジェクト設定 |

## セットアップ

### 1. スプレッドシート

- 使用するスプレッドシートのIDは `Code.js` の `SPREADSHEET_ID` で指定されています（仕様書のIDが既定値）。
- **スタッフDB** シート: 列 `uuid`, `name`, `birthdate`, `property`
- **打刻記録** シート: 列 `タイムスタンプ`, `スタッフID`, `氏名`, `出勤時刻`, `退勤時刻`, `休憩時間`
- 各スタッフ用シート: シート名は氏名の「姓」（先頭の単語）。列 `日付`, `始業`, `就業`, `休憩時間` など（1–2行目はヘッダー）。

### 2. GAS デプロイ

1. [Google Apps Script](https://script.google.com/) で新規プロジェクトを作成。
2. このリポジトリの `Code.js` と `index.html` の内容をコピーするか、clasp でプッシュ。
3. `.clasp.json` の `scriptId` に上記プロジェクトのスクリプトIDを設定し、`clasp push` で反映。
4. 「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」、実行ユーザー「自分」、アクセス「全員」でデプロイ。
5. 表示された「ウェブアプリのURL」を控える。

### 3. フロントの利用方法

- **GAS から配信**: 上記ウェブアプリのURLを**ブラウザのアドレスバーで直接開いて**ください。GAS_URL は自動で挿入されます。  
  **注意**: エディタの「プレビュー」や iframe 内ではカメラが許可されず QR が使えません。必ず「デプロイ」→「ウェブアプリのURL」を新しいタブで開いて利用してください。
- **GitHub Pages**: `standalone.html` をデプロイする際、GAS の URL を次のいずれかで設定します。
  1. **推奨**: `standalone.html` 内の `GAS_URL` の既定値（`return 'YOUR_GAS_WEB_APP_URL'` の部分）を、手順2で控えたウェブアプリのURLに書き換えてから push。
  2. ページを開くときの URL に `?gasUrl=https://script.google.com/macros/s/.../exec` を付ける（ブックマークやリンクで利用可）。
  3. `config.example.js` を `config.js` にコピーし URL を記入。`standalone.html` の jsQR の直前に行を追加: `<script src="config.js"></script>`。`config.js` を `.gitignore` に追加すると URL をリポジトリに含めずに済みます。

## 動作フロー

1. ページ読み込み時に `getStaffList` でスタッフ一覧を取得し、誕生日4桁（MMdd）ごとの候補マップをフロントで保持。
2. ユーザーが4桁入力すると、該当するスタッフのみ即表示（API 不要）。
3. スタッフを選択すると、属性に応じて勤務パターン（社員: 11-24 or 18-24、契約社員: 18-24）またはイレギュラー入力欄を表示。
4. QRコードをスキャンすると `recordTimestamp` を呼び出し、打刻記録シートとスタッフ別シートに追記。業務委託の場合は記録せず「記録なし」のみ返却。

## ライセンス・由来

フォーク元: [internQrAccessControl](https://github.com/hiroshimaeasyryo/internQrAccessControl)（入退室管理）。本プロジェクトは出勤管理用に仕様を変更しています。
