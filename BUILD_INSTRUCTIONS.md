# 教育計画PDFマージシステム v3.4.0 - ビルド手順

## 📋 概要

このドキュメントでは、教育計画PDFマージシステム v3.4.0 のインストーラーをビルドする手順を説明します。

## 🆕 v3.4.0 の主な変更点

### 初回セットアップエクスペリエンスの実装
- **初回セットアップウィザード**: 5ステップのガイド付きセットアップ
  - 年度設定（自動推定）
  - 作業フォルダ設定（パス検証付き）
  - Ghostscript自動検出
  - 設定完了サマリー

### 新規モジュール
- **ghostscript_detector.py** (260行) - Ghostscript自動検出
  - Windowsレジストリ検索
  - 環境変数チェック
  - 標準パス検索

- **config_validator.py** (257行) - 設定検証システム
  - 3段階の検証レベル（ERROR/WARNING/INFO）
  - 必須項目チェック
  - パス存在確認

- **gui/setup_wizard.py** (650+行) - セットアップウィザードUI
  - モーダルダイアログ
  - プログレスバー表示
  - リアルタイム検証

### 設定の改善
- config.jsonのテンプレート化（初回起動時にウィザードで設定）
- user_config.json（%LOCALAPPDATA%）での設定保存

## 🔧 必要な環境

- **Python**: 3.9 以上
- **OS**: Windows 10/11
- **必要なソフトウェア**:
  - Microsoft Office (Word/Excel/PowerPoint)
  - 一太郎（一太郎ファイル変換を使用する場合）

## 📦 ビルド手順

### 1. リポジトリのクローン／ダウンロード

```bash
# Git でクローン
git clone <repository-url>
cd <project-directory>

# または ZIP でダウンロードして解凍
```

### 2. 仮想環境の作成とアクティベート

```bash
# 仮想環境作成
python -m venv .venv

# アクティベート（Windows）
.venv\Scripts\activate

# または
.venv\Scripts\activate.bat
```

### 3. 依存パッケージのインストール

```bash
# requirements.txt から一括インストール
pip install -r requirements.txt

# PyInstaller もインストール
pip install pyinstaller
```

### 4. ビルド実行

#### 方法1: バッチファイルを使用（推奨）

```bash
# build.bat をダブルクリック、または：
build.bat
```

このスクリプトは以下を自動実行します：
1. クリーンアップ（build/, dist/ の削除）
2. 構文チェック（新規追加）
3. PyInstaller でのビルド
4. config.json のコピー

#### 方法2: 手動で実行

```bash
# クリーンアップ
rmdir /s /q build dist

# 構文チェック
python -m py_compile pdf_converter.py converters\office_converter.py converters\image_converter.py converters\ichitaro_converter.py

# ビルド
pyinstaller build_installer.spec --clean

# config.json をコピー
copy config.json dist\config.json
```

### 5. ビルド結果の確認

ビルドが成功すると、以下のファイルが生成されます：

```
dist/
  ├── 教育計画PDFマージシステム.exe  ← メインの実行ファイル
  └── config.json                      ← 設定ファイル
```

## 🧪 動作確認

### 1. テスト実行

```bash
# dist フォルダに移動
cd dist

# 実行ファイルを起動
教育計画PDFマージシステム.exe
```

### 2. 確認項目

- [ ] アプリケーションが起動する
- [ ] GUIが正しく表示される
- [ ] バージョンが 3.4.0 と表示される
- [ ] 初回起動時（user_config.json がない場合）にセットアップウィザードが表示される
- [ ] 各タブ（PDF、Excel、設定、ファイル）が機能する
- [ ] ログが正しく表示される（converters モジュールのログを含む）

## 📝 ビルド設定ファイル

### build_installer.spec

PyInstaller のビルド設定ファイル。以下の重要な設定が含まれます：

```python
hiddenimports = [
    # 新規追加（v3.4.0）
    'config_validator',
    'ghostscript_detector',
    'gui.setup_wizard',

    # v3.3.1で追加
    'pdf_merge_orchestrator',

    # v3.3.0で追加
    'converters',
    'converters.office_converter',
    'converters.image_converter',
    'converters.ichitaro_converter',

    # 既存モジュール
    'pdf_converter',
    'pdf_processor',
    'document_collector',
    # ... その他
]
```

### version_info.txt

Windows 実行ファイルのバージョン情報：

```
FileVersion: 3.4.0
ProductVersion: 3.4.0
ProductName: 教育計画PDFマージシステム
```

## 🔍 トラブルシューティング

### ビルドエラー: モジュールが見つからない

**症状**: `ModuleNotFoundError: No module named 'converters'`

**解決策**:
1. `build_installer.spec` の `hiddenimports` に converters モジュールが含まれているか確認
2. `converters/__init__.py` が存在するか確認
3. クリーンビルドを実行: `build.bat`

### 実行時エラー: ログが表示されない

**症状**: GUI でログが表示されない

**解決策**:
1. `gui/tabs/base_tab.py` のロガー名リストを確認
2. 以下のロガーが含まれているか確認:
   - `'converters.office_converter'`
   - `'converters.image_converter'`
   - `'converters.ichitaro_converter'`

### ビルドファイルサイズが大きい

**症状**: .exe ファイルが 100MB を超える

**解決策**:
- 正常です。以下のライブラリが含まれるため：
  - Python インタープリタ
  - PyQt/customtkinter
  - PyPDF2, PyMuPDF, reportlab
  - PIL/Pillow
  - pywin32, pywinauto

## 📚 参考情報

### プロジェクト構成（v3.4.0）

```
c:\Projects/
├── pdf_converter.py              ← ファサード（151行）
├── converters/                   ← v3.3.0で追加
│   ├── __init__.py
│   ├── office_converter.py       ← Office変換（233行）
│   ├── image_converter.py        ← 画像変換（48行）
│   └── ichitaro_converter.py     ← 一太郎変換（612行）
├── pdf_processor.py
├── pdf_merge_orchestrator.py     ← v3.3.1で追加
├── document_collector.py
├── config_loader.py
├── config_validator.py           ← v3.4.0で追加
├── ghostscript_detector.py       ← v3.4.0で追加
├── constants.py                  ← VERSION = "3.4.0"
├── exceptions.py
├── path_validator.py
├── folder_structure_detector.py
├── logging_config.py
├── update_excel_files.py
├── run_app.py
├── gui/
│   ├── app.py
│   └── tabs/
│       ├── base_tab.py           ← converters ロガー対応
│       ├── pdf_tab.py
│       ├── excel_tab.py
│       ├── settings_tab.py
│       └── file_tab.py
├── config.json
├── requirements.txt
├── build_installer.spec          ← PyInstaller設定
├── version_info.txt              ← バージョン情報
├── build.bat                     ← ビルドスクリプト
└── BUILD_INSTRUCTIONS.md         ← このファイル
```

### 変更履歴

#### v3.4.0 (2026-01-16)
- **初回セットアップウィザード**: 5ステップのガイド付きセットアップ
- **Ghostscript自動検出**: レジストリ・環境変数・標準パスの検索
- **設定検証システム**: ERROR/WARNING/INFOの3段階検証
- **config.jsonのテンプレート化**: 汎用的な設定ファイル

#### v3.3.1 (2026-01-15)
- リファクタリング漏れの修正
- 例外処理の統一化
- 定数管理の改善

#### v3.3.0 (2025-01-14)
- **重大なリファクタリング**: PDF変換モジュールの分割
- 単一責任の原則に準拠
- GUIログ統合の改善
- 定数の整理と明確化
- 100% docstring カバレッジ

#### v3.2.4 (以前)
- Excel転記機能
- フォルダ構造自動検出
- 一太郎変換機能

## 📞 サポート

問題が発生した場合は、以下の情報を含めて報告してください：

1. Python バージョン: `python --version`
2. OS バージョン: Windows 10/11
3. エラーメッセージ（全文）
4. ビルド手順（どこで失敗したか）

## 📄 ライセンス

このプロジェクトは内部使用を目的としています。
