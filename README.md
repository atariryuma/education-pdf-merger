# 教育計画PDFマージシステム

[![Python Version](https://img.shields.io/badge/python-3.8%2B-blue)](https://www.python.org/downloads/)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)](https://www.microsoft.com/windows)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)

教育計画・行事計画関連のドキュメントをPDFに変換・マージするWindowsアプリケーション。Word、Excel、PowerPoint、一太郎、画像などの複数形式に対応し、自動的に目次とブックマークを生成します。

## ✨ 主な機能

### PDF変換・マージ
- 📄 **複数形式対応**: Word (.doc/.docx)、Excel (.xls/.xlsx)、PowerPoint (.ppt/.pptx)、一太郎 (.jtd)、画像 (.jpg/.png/.bmp/.tiff)、PDF
- 📑 **自動目次生成**: ディレクトリ構造から階層的な目次を自動作成
- 🔖 **PDFブックマーク**: クリック可能なしおり（アウトライン）を自動設定
- 🔢 **ページ番号付加**: 表紙を除いて自動的にページ番号を挿入
- 🗜️ **PDF圧縮**: Ghostscriptによる自動圧縮でファイルサイズを最大70%削減
- 📋 **区切りページ**: フォルダごとに見出しページを自動挿入

### Excel処理
- 🔄 **自動データ反映**: 年間行事計画から様式4へのデータ自動転記
- 🧮 **時数自動計算**: 学年別の行事時数・欠時数を自動カウント
- 🗑️ **不要シート削除**: PDCAファイルから指定シートを一括削除

### ファイル管理
- 🏷️ **ファイル名整理**: PDCA優先の自動連番付与
- 📂 **超長パス対応**: 260文字を超えるパスのファイルを強制削除

## 📦 インストール

### 前提条件

- **Python 3.8以上** - [ダウンロード](https://www.python.org/downloads/)
- **Microsoft Office** - Word、Excel、PowerPointのインストールが必要
- **Ghostscript** - [ダウンロード](https://ghostscript.com/releases/gsdnld.html)
- **一太郎** (オプション) - .jtdファイルを扱う場合のみ

### 方法1: pipでインストール（開発者向け）

```bash
# リポジトリをクローン
git clone https://github.com/yourusername/pdf-merge-system.git
cd pdf-merge-system

# 依存関係をインストール
pip install -r requirements.txt

# または開発用の依存関係も含めてインストール
pip install -r requirements-dev.txt
```

### 方法2: インストーラーを使用（推奨）

1. [Releases](https://github.com/yourusername/pdf-merge-system/releases) から最新の `PDFMergeSystem_Setup_x.x.exe` をダウンロード
2. インストーラーを実行
3. デスクトップアイコンから起動

## 🚀 使い方

### GUIアプリケーション（推奨）

```bash
python run_app.py
```

または、インストーラー版の場合はデスクトップアイコンをダブルクリック。

#### 基本的な流れ

1. **「PDF統合」タブ**
   - 入力ディレクトリを選択（変換したいファイルがあるフォルダ）
   - 出力PDFファイル名を指定
   - 教育計画 or 行事計画を選択
   - 「PDF統合を実行」ボタンをクリック

2. **「Excel処理」タブ**
   - 対象のExcelファイルを開いた状態で実行
   - 「Excelデータ更新を実行」ボタンをクリック

3. **「設定」タブ**
   - 年度情報を更新（例: 令和８年度(2026)、R8）
   - パス設定を確認・変更
   - 「💾 設定を保存」ボタンで保存

### コマンドライン（上級者向け）

```bash
# 教育計画PDFの作成
python convert_and_merge.py
```

**注意**: 以前のバージョンにあった以下のスクリプトは削除されました（GUIアプリに統合）:
- `convert_and_merge_event.py` - 行事計画PDF作成（GUIの「PDF統合」タブで実行可能）
- `update_excel_files.py` - Excelデータ更新（GUIの「Excel処理」タブで実行可能）
- `rename_file.py` - ファイル名整理（GUIの「ファイル管理」タブで実行可能）
- `delete.py` - 不要シート削除（GUIの「ファイル管理」タブで実行可能）
- `force_delete.py` - 長いパスの削除ユーティリティ（不要になったため削除）

## ⚙️ 設定

### config.json の編集

プロジェクトルートの `config.json` で以下の設定が可能です：

```json
{
  "year": "令和７年度(2025)",
  "year_short": "R7",
  "base_paths": {
    "google_drive": "G:\\マイドライブ\\ドキュメント",
    "network": "\\\\10.206.2.16\\天久小\\share",
    "local_temp": "C:\\Projects\\temp_pdfs"
  },
  "ghostscript": {
    "executable": "C:\\Program Files\\gs\\gs10.04.0\\bin\\gswin64c.exe"
  }
}
```

**主な設定項目:**

| 項目 | 説明 | 例 |
|------|------|-----|
| `year` | 年度（フル） | `"令和８年度(2026)"` |
| `year_short` | 年度（略称） | `"R8"` |
| `base_paths.google_drive` | Google Driveルートパス | `"G:\\マイドライブ"` |
| `base_paths.network` | ネットワークパス | `"\\\\server\\share"` |
| `ghostscript.executable` | Ghostscript実行ファイルパス | `"C:\\Program Files\\gs\\...\\gswin64c.exe"` |

GUIアプリの「設定」タブから直接編集することも可能です。

## 📁 プロジェクト構成

```
pdf-merge-system/
├── config.json                 # 設定ファイル
├── pyproject.toml              # プロジェクトメタデータ
├── requirements.txt            # 依存関係
├── requirements-dev.txt        # 開発用依存関係
│
├── run_app.py                  # GUIアプリケーション起動
├── convert_and_merge.py        # 教育計画PDF作成
│
├── config_loader.py            # 設定読み込み
├── pdf_converter.py            # PDF変換エンジン
├── pdf_processor.py            # PDF処理（マージ、圧縮など）
├── document_collector.py       # ドキュメント収集・目次生成
├── ghostscript_utils.py        # Ghostscript操作
├── logging_config.py           # ロギング設定
├── exceptions.py               # 例外クラス
├── constants.py                # 定数定義
│
├── gui/                        # GUIモジュール
├── tests/                      # テストコード
├── installer/                  # インストーラー設定
└── temp_pdfs/                  # 一時ファイル
```

## 🏗️ アーキテクチャ

### 処理フロー

```
設定読み込み (config.json)
    ↓
ディレクトリ探索 & ファイル収集
    ↓
各ファイルをPDFに変換
 ├─ Office文書 → Win32 COM
 ├─ 一太郎 → pywinauto
 ├─ 画像 → reportlab
 └─ PDF → パススルー
    ↓
一時的にマージ
    ↓
目次PDFを生成
    ↓
表紙・目次・本文を結合
    ↓
ページ番号を追加
    ↓
しおり（アウトライン）を設定
    ↓
Ghostscriptで圧縮
    ↓
最終PDF完成 ✓
```

### モジュール構成

- **config_loader.py**: 設定ファイルの読み込みとパス構築
- **pdf_converter.py**: 各種形式からPDFへの変換
- **pdf_processor.py**: PDFのマージ、圧縮、ページ番号追加、ブックマーク設定
- **document_collector.py**: ディレクトリ探索と目次生成、全体の処理フロー制御

## 🔧 開発

### 開発環境のセットアップ

```bash
# 仮想環境の作成
python -m venv venv

# 仮想環境の有効化
venv\Scripts\activate  # Windows

# 開発用依存関係のインストール
pip install -r requirements-dev.txt
```

### テストの実行

```bash
# すべてのテストを実行
pytest

# カバレッジ付きで実行
pytest --cov=. --cov-report=html
```

### コード品質チェック

```bash
# フォーマット（自動修正）
black .

# Lint
flake8 .

# 型チェック
mypy .
```

### EXEファイルのビルド

```bash
# ビルドスクリプトを実行
build.bat

# 出力先: dist\教育計画PDFマージシステム.exe
```

インストーラーの作成:
```bash
cd installer
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" setup.iss
```

## 📄 ライセンス

MIT License - 詳細は [LICENSE](LICENSE) をご覧ください。

## 🤝 コントリビューション

プルリクエストを歓迎します！大きな変更の場合は、まずissueを開いて変更内容を議論してください。

### コントリビューションの手順

1. このリポジトリをフォーク
2. フィーチャーブランチを作成 (`git checkout -b feature/AmazingFeature`)
3. 変更をコミット (`git commit -m 'Add some AmazingFeature'`)
4. ブランチにプッシュ (`git push origin feature/AmazingFeature`)
5. プルリクエストを開く

## ⚠️ トラブルシューティング

### エラー: "config.json not found"
**原因**: 設定ファイルが見つからない
**解決**: `config.json` がスクリプトと同じディレクトリにあることを確認

### エラー: "Invalid JSON"
**原因**: config.jsonの構文エラー
**解決**: [JSONLint](https://jsonlint.com/) で検証、カンマや引用符を確認

### PDF変換が失敗する
**原因**: Office/一太郎/Ghostscriptがインストールされていない
**解決**: 必要なソフトウェアをインストール、パスを確認

### 一太郎変換が途中で止まる
**原因**: UIオートメーションのタイミング問題
**解決**: `config.json` の一太郎設定でタイミングを調整

### Excelスクリプトが動かない
**原因**: ファイルが開かれていない、ファイル名不一致
**解決**: 対象Excelファイルを開く、ファイル名を確認

### 変換が遅い
**原因**: Ghostscript圧縮は時間がかかる
**対策**: 小さいファイルで試す、圧縮をスキップ（開発時）

## 📚 参考資料

- [PyPDF2 ドキュメント](https://pypdf2.readthedocs.io/)
- [ReportLab ユーザーガイド](https://www.reportlab.com/docs/reportlab-userguide.pdf)
- [pywin32 ドキュメント](https://github.com/mhammond/pywin32)
- [pywinauto ドキュメント](https://pywinauto.readthedocs.io/)

## 📝 変更履歴

詳細は [CHANGELOG.md](CHANGELOG.md) をご覧ください。

### v3.2 (2025-12-15)
- ✨ ベストプラクティスに準拠（pyproject.toml、.gitignore追加）
- 🚀 一太郎変換処理を40%高速化（22秒 → 12-15秒）
- 🎯 pywinauto ベストプラクティス適用（window.close()、app.kill()）
- 🧹 プロジェクトクリーンアップ（不要ファイル削除、~115MB削減）

### v3.0 (2025-11-26)
- 🎨 GUIアプリケーション追加
- ⚙️ 設定タブで年度情報を直接編集可能に
- 🔄 設定の保存・再読み込み機能

### v2.0 (2025-11)
- 🏗️ モジュール化によるリファクタリング
- 📦 設定ファイル方式の導入
- 🔧 年度更新の簡素化

## 👥 作者

**School Tools**

## 🙋 サポート

問題が発生した場合は、[GitHub Issues](https://github.com/yourusername/pdf-merge-system/issues) で報告してください。

報告時は以下の情報を含めてください：
- エラーメッセージ（完全な内容）
- 実行したコマンド
- `config.json` の内容（機密情報は削除）
- Python、Office、Ghostscriptのバージョン

---

**Made with ❤️ for Education**
