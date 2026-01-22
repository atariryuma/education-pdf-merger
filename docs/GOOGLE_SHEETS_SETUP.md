# Google Sheets連携セットアップガイド

教育計画PDFマージシステムでGoogle Sheetsを参照元として使用するための設定手順です。

## 概要

Google Sheetsをデータ参照元として使用することで、以下のメリットがあります:

- クラウド上でデータを一元管理
- 複数人での同時編集が可能
- どこからでもアクセス可能
- 自動バックアップ

## 前提条件

- Googleアカウント（個人・組織アカウント両対応）
- インターネット接続
- Google Sheetsの参照元スプレッドシート

---

## ステップ1: Google Cloud プロジェクトの作成

### 1.1 Google Cloud Consoleにアクセス

https://console.cloud.google.com/ にアクセスしてログイン

### 1.2 新規プロジェクトを作成

1. 画面上部の「プロジェクトを選択」をクリック
2. 「新しいプロジェクト」をクリック
3. プロジェクト名: `education-pdf-merger`（任意）
4. 「作成」をクリック

---

## ステップ2: Google Sheets API の有効化

### 2.1 APIライブラリにアクセス

1. 左メニュー「APIとサービス」→「ライブラリ」
2. 検索バーに「Google Sheets API」と入力
3. 「Google Sheets API」をクリック
4. 「有効にする」をクリック

---

## ステップ3: OAuth 2.0 認証情報の作成

### 3.1 認証情報の作成

1. 左メニュー「APIとサービス」→「認証情報」
2. 「認証情報を作成」→「OAuth クライアント ID」をクリック

### 3.2 同意画面の設定（初回のみ）

OAuth同意画面の設定が求められた場合:

1. 「同意画面を構成」をクリック
2. User Type: **外部** を選択（組織アカウントの場合は「内部」も可）
3. 「作成」をクリック
4. アプリ情報:
   - アプリ名: `教育計画PDFマージシステム`
   - ユーザーサポートメール: 自分のメールアドレス
   - デベロッパーの連絡先情報: 自分のメールアドレス
5. 「保存して次へ」をクリック
6. スコープ: **何も追加せず**「保存して次へ」
7. テストユーザー: **追加不要**「保存して次へ」
8. 「ダッシュボードに戻る」をクリック

### 3.3 OAuth クライアント ID の作成

1. 再度「認証情報を作成」→「OAuth クライアント ID」
2. アプリケーションの種類: **デスクトップ アプリ**
3. 名前: `教育計画PDFマージ`（任意）
4. 「作成」をクリック

### 3.4 認証情報のダウンロード

1. 作成完了ダイアログが表示されたら「JSONをダウンロード」をクリック
2. ダウンロードしたファイル名を `credentials.json` に変更

---

## ステップ4: credentials.json の配置

ダウンロードした `credentials.json` を以下のいずれかの場所に配置:

### オプション1: アプリケーションディレクトリ（推奨）

```
c:\Projects\education-pdf-merger\credentials.json
```

### オプション2: ユーザーデータディレクトリ

```
%LOCALAPPDATA%\PDFMergeSystem\credentials.json
```

通常のパス例:
```
C:\Users\[ユーザー名]\AppData\Local\PDFMergeSystem\credentials.json
```

---

## ステップ5: Google Sheets の準備

### 5.1 スプレッドシートの要件

参照元のGoogle Sheetsは以下の構造が必要です:

- **シート名**: `メインデータ`（設定で変更可能）
- **C列**: 行事名（検索キー）
- **A列**: 日付
- **E～AN列**: 6学年×6校時の行事データ（36列）

### 5.2 URLの取得

1. Google Sheetsでスプレッドシートを開く
2. アドレスバーからURLをコピー
3. URL形式: `https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit#gid=0`

---

## ステップ6: アプリケーション設定

### 6.1 config.json の編集

`config.json`（または設定タブ）で以下を設定:

```json
{
  "files": {
    "reference_mode": "google_sheets",
    "google_sheets_reference_url": "https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/edit",
    "google_sheets_reference_sheet": "メインデータ",
    "excel_target": "様式4-3.xlsx"
  }
}
```

**重要**: `excel_target`（ターゲットExcelファイル）は引き続き必要です。

---

## ステップ7: 初回認証

### 7.1 アプリケーション起動

1. アプリケーションを起動
2. 「📊 Excel処理」タブを開く
3. 「▶ Excelデータ更新を実行」をクリック

### 7.2 OAuth認証フロー

初回実行時、ブラウザが自動的に開きます:

1. Googleアカウントでログイン
2. 「このアプリは Google で確認されていません」警告が表示される場合:
   - 「詳細」をクリック
   - 「[アプリ名]（安全ではないページ）に移動」をクリック
3. 「Google Sheets API への読み取り専用アクセスを許可」で「許可」をクリック
4. 「認証が完了しました」と表示されたらブラウザを閉じる

### 7.3 認証情報の保存

認証が完了すると、トークンが以下に保存されます:

```
%LOCALAPPDATA%\PDFMergeSystem\google_credentials\token.json
```

次回以降は自動的にこのトークンが使用され、ブラウザは開きません。

---

## トラブルシューティング

### ❌ エラー: 認証情報ファイルが見つかりません

**原因**: `credentials.json` が正しい場所にない

**解決策**:
1. `credentials.json` の配置場所を確認
2. ファイル名が正確に `credentials.json` であることを確認

### ❌ エラー: Google Sheets URLが無効です

**原因**: URL形式が正しくない

**解決策**:
- URL形式: `https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/...`
- `/spreadsheets/d/` の直後がスプレッドシートIDです

### ❌ エラー: シートが見つかりません

**原因**: シート名が一致しない

**解決策**:
1. Google Sheetsで実際のシート名を確認
2. `config.json` の `google_sheets_reference_sheet` を正しいシート名に変更

### ❌ エラー: API利用上限に達しました

**原因**: Google Sheets APIの利用制限（1分間に60リクエスト）

**解決策**:
- 数分待ってから再実行
- 通常の使用では制限に達することはありません

### ⚠️ ファイアウォールでブラウザが開かない

**原因**: 企業ネットワークでlocalhost通信がブロックされている

**解決策**:
1. IT部門にlocalhost:8080-8090の通信許可を依頼
2. または、個人ネットワークで初回認証を実施

---

## セキュリティに関する注意事項

### ✅ 安全な設計

- **読み取り専用**: Google Sheets APIは読み取り専用スコープのみ使用
- **ローカル認証**: トークンはローカルマシンにのみ保存
- **暗号化通信**: すべての通信はHTTPS経由

### ⚠️ 注意点

- `credentials.json`と`token.json`は**機密情報**です
- 他人と共有しないでください
- Gitリポジトリにコミットしないでください

### 🔒 認証解除

認証を解除したい場合:

1. 以下のファイルを削除:
   ```
   %LOCALAPPDATA%\PDFMergeSystem\google_credentials\token.json
   ```
2. 次回実行時に再度OAuth認証が必要になります

---

## よくある質問

### Q: 組織アカウントで使用できますか？

**A**: はい。Google Workspace（旧G Suite）アカウントでも使用可能です。IT管理者がGoogle Sheets APIを無効にしていない限り利用できます。

### Q: オフラインで使用できますか？

**A**: いいえ。Google Sheetsへのアクセスにはインターネット接続が必要です。

### Q: 複数のスプレッドシートを切り替えられますか？

**A**: はい。`config.json`の`google_sheets_reference_url`を変更するだけで切り替わります。

### Q: Excelモードに戻せますか？

**A**: はい。`config.json`の`reference_mode`を`"excel"`に変更してください。

---

## サポート

問題が解決しない場合:

1. ログファイルを確認: アプリケーションディレクトリの `logs/` フォルダ
2. GitHubでIssueを作成: https://github.com/anthropics/claude-code/issues

---

**最終更新**: 2026-01-22
**バージョン**: 1.0.0
