# インストーラー作成手順

## 前提条件

### 1. Inno Setup 6 のインストール

1. 以下のURLからInno Setup 6をダウンロード:
   - https://jrsoftware.org/isdl.php
   - 「Inno Setup 6.x.x」の「Download」をクリック

2. ダウンロードしたインストーラーを実行
   - 日本語言語ファイルを含めてインストール

### 2. EXEのビルド

プロジェクトルートで以下を実行:
```cmd
cd c:\Projects
build.bat
```

## インストーラーのビルド

```cmd
cd c:\Projects\installer
build_installer.bat
```

## 出力

成功すると以下にインストーラーが作成されます:
```
c:\Projects\dist\installer\PDFMergeSystem_Setup_3.1.exe
```

## インストーラーの機能

- **インストール先選択**: ユーザーが指定可能
- **Ghostscriptパス設定**: インストール中に設定
- **Google Driveパス設定**: インストール中に設定
- **デスクトップショートカット**: オプション
- **スタートメニュー登録**: 自動
- **アンインストーラー**: 自動生成

## ファイル構成

```
installer/
├── setup.iss              # Inno Setupスクリプト
├── config_template.json   # 設定ファイルテンプレート
├── build_installer.bat    # ビルドスクリプト
└── README.md              # この文書
```

## カスタマイズ

### バージョン変更
`setup.iss`の以下を編集:
```
#define MyAppVersion "3.1"
```

### アイコン追加
`setup.iss`の[Setup]セクションに追加:
```
SetupIconFile=app_icon.ico
```
