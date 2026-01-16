# インストーラービルド手順

## 📋 概要

このディレクトリには、教育計画PDFマージシステム v3.4.0 のWindowsインストーラーを作成するためのファイルが含まれています。

## 🔧 必要なソフトウェア

1. **Inno Setup 6.x 以降**
   - ダウンロード: https://jrsoftware.org/isdl.php
   - インストール先: `C:\Program Files (x86)\Inno Setup 6\` (推奨)

2. **ビルド済みの実行ファイル**
   - `dist\教育計画PDFマージシステム.exe`
   - 先に `build.bat` を実行してビルドしておく必要があります

## 📁 ファイル構成

```
installer/
  ├── setup.iss                  ← Inno Setup スクリプト
  ├── build_installer.bat        ← ビルドスクリプト
  ├── README_INSTALLER.md        ← このファイル
  └── output/                    ← 生成されたインストーラー（ビルド後）
      └── PDFMergeSystem_Setup_3.4.0.exe
```

## 🚀 インストーラーのビルド手順

### ステップ1: アプリケーションのビルド

インストーラーをビルドする前に、まずアプリケーション本体をビルドします。

```bash
# プロジェクトルートで実行
cd c:\Projects
build.bat
```

ビルドが成功すると、以下のファイルが生成されます：
- `dist\教育計画PDFマージシステム.exe`
- `dist\config.json`

### ステップ2: インストーラーのビルド

```bash
# installer ディレクトリに移動
cd installer

# ビルドスクリプト実行
build_installer.bat
```

または、Inno Setup を直接使用：

```bash
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" setup.iss
```

### ステップ3: 出力の確認

ビルドが成功すると、以下のファイルが生成されます：

```
dist\installer\
  └── PDFMergeSystem_Setup_3.4.0.exe  (約 100-150 MB)
```

## 📦 インストーラーに含まれるファイル

### 実行ファイル
- `教育計画PDFマージシステム.exe` - メイン実行ファイル

### 設定ファイル
- `config.json` - アプリケーション設定

### ドキュメント (docs/ フォルダ)
- `CHANGELOG.txt` - 変更履歴
- `BUILD_INSTRUCTIONS.txt` - ビルド手順
- `RELEASE_NOTES.txt` - v3.4.0 リリースノート

## 🎯 インストーラーの機能

### インストール時
1. **前提条件チェック**
   - アプリケーションが実行中の場合は自動終了

2. **ファイル配置**
   - `C:\Program Files\教育計画PDFマージシステム\` にインストール
   - ドキュメントを `docs\` サブフォルダに配置

3. **ディレクトリ作成**
   - ログディレクトリ: `%LOCALAPPDATA%\PDFMergeSystem\logs`
   - 一時ファイルディレクトリ: `%LOCALAPPDATA%\PDFMergeSystem\temp`

4. **スタートメニュー**
   - アプリケーション起動
   - 設定ファイル
   - ドキュメント
   - リリースノート v3.4.0
   - アンインストール

5. **オプション**
   - デスクトップアイコン作成（デフォルト: 無効）

### アンインストール時
1. **プロセス終了**
   - アプリケーションが実行中の場合は自動終了

2. **ファイル削除**
   - インストールディレクトリ
   - ユーザーデータディレクトリ
   - ログファイル
   - 一時ファイル

3. **完全なクリーンアップ**
   - すべての設定とログを削除

## ⚙️ setup.iss のカスタマイズ

### バージョン情報の更新

```pascal
#define MyAppName "教育計画PDFマージシステム"
#define MyAppVersion "3.4.0"         ← ここを更新
#define MyAppPublisher "教育機関向けPDFツール"
```

### インストール先の変更

```pascal
[Setup]
DefaultDirName={autopf}\{#MyAppName}  ← デフォルトインストール先
```

### 含めるファイルの追加

```pascal
[Files]
Source: "..\dist\新しいファイル.dll"; DestDir: "{app}"; Flags: ignoreversion
```

### スタートメニュー項目の追加

```pascal
[Icons]
Name: "{group}\新しい項目"; Filename: "{app}\新しいファイル.exe"
```

## 🐛 トラブルシューティング

### エラー: Inno Setup が見つかりません

**症状**: `Inno Setup 6 が見つかりません`

**解決策**:
1. Inno Setup 6.x をインストール: https://jrsoftware.org/isdl.php
2. インストール先を確認: `C:\Program Files (x86)\Inno Setup 6\`
3. カスタムパスの場合は `build_installer.bat` の `ISCC` 変数を更新

### エラー: EXE ファイルが見つかりません

**症状**: `EXEファイルが見つかりません`

**解決策**:
1. プロジェクトルートで `build.bat` を実行
2. `dist\教育計画PDFマージシステム.exe` が存在することを確認

### エラー: ドキュメントが見つかりません

**症状**: `CHANGELOG.md が見つかりません`

**解決策**:
- プロジェクトルートにドキュメントファイルがあることを確認
- ビルドは警告付きで続行されます

### ビルドは成功するがインストーラーが起動しない

**症状**: インストーラーをダブルクリックしても起動しない

**解決策**:
1. 管理者権限で実行してみる
2. ウイルス対策ソフトが警告を出している可能性を確認
3. ファイルサイズを確認（正常: 100-150 MB程度）

## 📝 ビルド後の確認項目

インストーラーのビルド後、以下を確認してください：

- [ ] インストーラーが生成されている
- [ ] ファイルサイズが適切（100-150 MB程度）
- [ ] テスト環境でインストールできる
- [ ] アプリケーションが正常に起動する
- [ ] バージョンが 3.4.0 と表示される
- [ ] 初回起動時にセットアップウィザードが表示される（user_config.json がない場合）
- [ ] すべてのドキュメントが含まれている
- [ ] スタートメニューに項目が作成される
- [ ] アンインストールが正常に動作する

## 🔗 参考リンク

- [Inno Setup 公式サイト](https://jrsoftware.org/isinfo.php)
- [Inno Setup ドキュメント](https://jrsoftware.org/ishelp/)
- [Inno Setup サンプルスクリプト](https://github.com/jrsoftware/issrc/tree/main/Examples)

## 📞 サポート

問題が発生した場合は、以下の情報を含めて報告してください：

1. エラーメッセージ（全文）
2. 実行したコマンド
3. Inno Setup のバージョン
4. Windows のバージョン
5. ビルド環境の詳細

---

**バージョン**: 3.4.0
**最終更新**: 2026-01-16
