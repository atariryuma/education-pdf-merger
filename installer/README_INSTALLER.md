# インストーラービルド手順

## 概要

教育計画PDFマージシステムの Windows インストーラー (`.exe`) を Inno Setup で作成する手順。

バージョンは `pyproject.toml` を単一の真実とし、`installer/setup.iss` と `version_info.txt`、`build.bat` を同期させる。

## 必要なソフトウェア

1. **Inno Setup 6.x 以降**
   - ダウンロード: <https://jrsoftware.org/isdl.php>
   - 推奨インストール先: `C:\Program Files (x86)\Inno Setup 6\`

2. **ビルド済みの実行ファイル**
   - `dist\教育計画PDFマージシステム.exe`
   - 先に `build.bat`（プロジェクトルート）を実行してビルド済みであること

## ファイル構成

```text
installer/
  ├─ setup.iss               Inno Setup スクリプト
  ├─ build_installer.bat     インストーラービルドスクリプト
  └─ README_INSTALLER.md     このファイル
```

生成物は `dist/installer/PDFMergeSystem_Setup_<version>.exe`。

## ビルド手順

### 1. アプリケーション本体のビルド

プロジェクトルートで実行:

```bat
build.bat
```

成功すると `dist\教育計画PDFマージシステム.exe` と `dist\config.json` が生成される。

### 2. インストーラーのビルド

```bat
cd installer
build_installer.bat
```

または Inno Setup の `ISCC.exe` を直接呼び出す:

```bat
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" setup.iss
```

### 3. 出力の確認

`dist\installer\PDFMergeSystem_Setup_<version>.exe`（100〜150 MB）が生成される。

## インストーラーの動作

### インストール時

- 実行中のアプリケーションを終了
- `C:\Program Files\教育計画PDFマージシステム\` にファイル配置
- `%LOCALAPPDATA%\PDFMergeSystem\logs` と `temp` を作成
- スタートメニューに登録（アプリ・設定ファイル・アンインストール）
- デスクトップアイコンはオプション（デフォルト無効）

### アンインストール時

- 実行中のアプリケーションを終了
- インストールディレクトリ・ユーザーデータ・ログをすべて削除

## バージョンアップ手順

1. `pyproject.toml` の `version` を更新（単一の真実）
2. 以下を同じバージョンに更新:
   - `shared/constants.py` の `AppConstants.VERSION`
   - `installer/setup.iss` の `MyAppVersion`
   - `installer/build_installer.bat` のエコー文
   - `version_info.txt` の `filevers` / `prodvers` / `FileVersion` / `ProductVersion`
   - `build_installer.spec` のヘッダーコメント
   - `build.bat` のエコー文
3. `CHANGELOG.md` に変更内容を追記
4. `build.bat` → `build_installer.bat` の順でビルド

## トラブルシューティング

| 症状 | 対策 |
| --- | --- |
| `Inno Setup 6 が見つかりません` | Inno Setup 6.x をインストール、または `build_installer.bat` の `ISCC` 変数を編集 |
| `EXEファイルが見つかりません` | プロジェクトルートで `build.bat` を先に実行 |
| インストーラーが起動しない | 管理者権限で実行、ウイルス対策ソフトの誤検知を確認 |

## 参考リンク

- [Inno Setup 公式](https://jrsoftware.org/isinfo.php)
- [Inno Setup ドキュメント](https://jrsoftware.org/ishelp/)
