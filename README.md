# 名刺自動処理システム

Google Driveにアップロードされた名刺画像を自動的に解析し、Google Spreadsheetsにデータを保存するシステムです。

## 機能

- 名刺画像の自動AI解析（Gemini API使用）
- Google Sheetsへの自動データ保存
- 重複ファイルの自動チェック
- 処理済みファイルの自動移動
- 30分間隔での自動実行

## 必要な設定

### 1. 環境変数の設定

`.env.example`をコピーして`.env`ファイルを作成し、以下の値を設定してください：

```bash
# Gemini API Key
GEMINI_API_KEY=your_gemini_api_key_here

# Google Service Account Key File Path
GOOGLE_SERVICE_ACCOUNT_KEY_FILE=path/to/your/service-account-key.json
```

### 2. Google Cloud Platform設定

1. Google Cloud Consoleでプロジェクトを作成
2. 以下のAPIを有効化：
   - Google Sheets API
   - Google Drive API
3. サービスアカウントを作成し、JSONキーファイルをダウンロード
4. Google SheetsとGoogle Driveへのアクセス権限を付与

### 3. Google Drive設定

指定されたフォルダID：
- 名刺フォルダ: `1Q29bAIoQ__PADA2NefymTpsOO_yAg9ee`
- 処理済みフォルダ: `16LAj4yAM2cyk-tUlY7WYFTkbmB5Krvjx`

### 4. Google Sheets設定

スプレッドシートのシート名: `名刺情報`

ヘッダー構成：
```
ファイルid | 名前 | フリガナ | 社名 | 社名フリガナ | 部署 | 肩書 | 郵便番号 | 住所
```

## インストール

```bash
npm install
```

## 使用方法

### 1. ビルド

```bash
npm run build
```

### 2. 実行

```bash
npm run business-card <スプレッドシートID>
```

例：
```bash
npm run business-card 1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms
```

## 処理フロー

1. 名刺フォルダ内の画像ファイルを取得
2. スプレッドシートの既存ファイルIDと照合
3. 新しいファイルのみを処理対象として選択
4. 各画像をGemini AIで解析
5. 解析結果をGoogle Sheetsに保存
6. 処理済みファイルを専用フォルダに移動
7. 30分後に再実行

## 抽出される情報

- ファイルID
- 名前
- フリガナ（カタカナ）
- 社名
- 社名フリガナ（カタカナ）
- 部署
- 肩書
- 郵便番号
- 住所

## 注意事項

- ファイル処理中にエラーが発生した場合、該当ファイルはスキップされ次のファイルが処理されます
- 重複チェックはファイルIDベースで行われます
- 処理済みファイルは移動（コピーではない）されます

主要機能

  1. 自動画像解析

  - Gemini AI APIによる名刺画像からの情報抽出
  - 抽出項目：
    - ファイルID、名前、フリガナ、社名、社名フリガナ
    - 部署、肩書、郵便番号、住所、メールアドレス

  2. データ管理

  - Google Sheetsへの自動保存
  - ヘッダー自動作成：シート「名刺情報」がない場合は新規
  作成
  - 列構成：A〜J列（ファイルID〜メールアドレス）

  3. 重複制御（2段階）

  - ファイルID重複チェック：既に処理済みのファイルをスキ
  ップ
  - 内容重複チェック：名前と会社名の完全一致で判定
    - 空白文字を除去して比較
    - 名前が一致した場合のみ会社名を確認（処理効率化）

  4. ファイル管理

  - 自動フォルダ移動：処理完了後、処理済みフォルダへ移動
  -
  重複ファイルも移動：データ未保存でも処理済みフォルダへ

  5. 定期実行

  - 30分間隔で自動実行
  - 名刺フォルダを定期監視

  処理フロー

  1. 名刺フォルダ内の全画像ファイルを取得
     ↓
  2. ファイルIDによる1次重複チェック
     ├─ 重複ファイル → 処理済みフォルダへ移動のみ
     └─ 新規ファイル → 次のステップへ
     ↓
  3. Gemini AIによる画像解析
     ↓
  4. 名前・会社名による2次重複チェック
     ├─ 重複データ → 保存スキップ
     └─ 新規データ → スプレッドシートに保存
     ↓
  5. 処理済みフォルダへファイル移動
     ↓
  6. 30分後に再実行

  使用技術

  - 言語：TypeScript/Node.js
  - API：Google Sheets API、Google Drive API、Gemini AI
  API
  - 認証：サービスアカウント認証

  フォルダ構成

  - 名刺フォルダID：1Q29bAIoQ__PADA2NefymTpsOO_yAg9ee
  -
  処理済みフォルダID：16LAj4yAM2cyk-tUlY7WYFTkbmB5Krvjx
  - スプレッドシートID：1_aS8cRFajNlrAnvSgjmNhhM8U8W8lmv
  bbBDxmbJ_w6Q

  運用上のメリット

  1. 完全自動化：手動作業不要
  2. 重複防止：ファイルと内容の2重チェック
  3. データ整理：処理済みファイルの自動整理
  4. 拡張性：AI解析項目の追加が容易