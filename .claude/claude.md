# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

> **最終更新**: 2026-01-26
> **バージョン**: 3.5.0

## プロジェクト概要

教育計画や行事計画のドキュメントをPDF化して統合するWindowsデスクトップアプリケーション。
Word、Excel、PowerPoint、一太郎、画像ファイルを目次付き単一PDFに変換します。

---

## 開発環境の前提条件

### 必須要件

- **OS**: Windows 10/11（Windowsのみ対応）
- **Python**: 3.8以上（推奨: 3.9+）
- **Microsoft Office**: Word、Excel、PowerPoint（COM経由で操作）
- **Ghostscript**: PDF圧縮用（自動検出可能）

### オプション

- **一太郎**: .jtdファイル変換を使用する場合のみ
- **Google Cloud プロジェクト**: Google Sheets参照機能を使用する場合（[詳細](docs/GOOGLE_SHEETS_SETUP.md)）

### 重要な制限事項

- ⚠️ **Windowsのみ**: Win32 COM、pywinautoを使用するためmacOS/Linuxは非対応
- ⚠️ **Office必須**: Word/Excel/PowerPoint変換にはOfficeインストールが必要
- ⚠️ **UI自動化**: 一太郎変換はUI自動化のため、実行中は他の操作を避ける

---

## 開発コマンド

### 環境セットアップ

```bash
# 仮想環境作成・有効化
python -m venv venv
venv\Scripts\activate

# 開発用依存関係インストール
pip install -r requirements-dev.txt

# pre-commitフック有効化
pre-commit install
```

### アプリケーション実行

```bash
# GUIアプリケーション起動（推奨）
python run_app.py

# pipインストール後のエントリーポイント
pdf-merge
```

### テスト

```bash
# 全テスト実行（カバレッジ付き）
pytest

# 特定テストファイル実行
pytest tests/test_pdf_converter.py

# マーカー指定実行
pytest -m unit           # ユニットテストのみ
pytest -m integration    # 統合テストのみ
pytest -m "not slow"     # 低速テストを除外

# 詳細出力
pytest -v --cov=. --cov-report=html --cov-report=term-missing
```

### リント・フォーマット

```bash
# Ruffでリント・フォーマット（高速）
ruff check --fix .
ruff format .

# 型チェック（厳密度はmypy.iniで制御）
mypy --config-file=mypy.ini

# セキュリティチェック
bandit -r . --skip B101,B601 --exclude tests/

# pre-commitフック手動実行
pre-commit run --all-files
```

### ビルド

```bash
# 実行ファイル生成（Windows専用）
build.bat
# または: pyinstaller build_installer.spec --clean

# 生成物: dist\教育計画PDFマージシステム.exe
```

---

## CI/CD

### GitHub Actions

3つのジョブで自動品質チェックを実施（`.github/workflows/ci.yml`）:

#### 1. **品質チェック** (quality)

```bash
# 型チェック
mypy --config-file mypy.ini --no-error-summary .

# リント
ruff check . --output-format=github

# セキュリティチェック
bandit -r . -f screen --skip B101,B601
```

#### 2. **テスト** (test)

```bash
# テスト実行 + カバレッジ
pytest tests/ --cov=. --cov-report=xml --cov-report=term-missing -v

# カバレッジレポートをCodecovにアップロード
```

#### 3. **ビルド検証** (build)

```bash
# PyInstallerでEXE生成
pyinstaller build_installer.spec --clean --noconfirm

# 生成物の検証
# 成果物を7日間保存
```

### トリガー

- `main`、`develop`ブランチへのpush
- `main`、`develop`へのプルリクエスト
- 手動実行（workflow_dispatch）

---

## アーキテクチャ概要

### レイヤー構造

```text
GUI層 (gui/)
  ↓ 依存
オーケストレーション層 (pdf_merge_orchestrator.py)
  ↓ 依存
ビジネスロジック層
  ├─ document_collector.py    (ディレクトリ走査・目次生成)
  ├─ pdf_converter.py          (形式変換ファサード)
  ├─ pdf_processor.py          (PDF操作)
  └─ Excel処理
      ├─ base_excel_transfer.py (抽象基底クラス)
      ├─ update_excel_files.py  (Excel COM実装)
      └─ google_sheets_transfer.py (Google Sheets API実装)
  ↓ 依存
インフラ層
  ├─ config_loader.py
  ├─ path_validator.py
  ├─ logging_config.py
  └─ converters/ (office_converter, image_converter, ichitaro_converter)
```

### 処理フロー（6ステップ）

`PDFMergeOrchestrator.create_merged_pdf()` が制御:

1. **ファイル収集・変換** → `DocumentCollector.collect_and_convert()`
2. **一時マージ** → `PDFProcessor.merge_pdfs()`
3. **目次PDF生成** → `PDFProcessor.create_toc_pdf()`
4. **表紙・目次・本文結合** → `PDFProcessor.merge_pdfs()`
5. **ページ番号追加** → `PDFProcessor.add_page_numbers()`
6. **ブックマーク設定** → `PDFProcessor.add_bookmarks()`

オプション: `GhostscriptCompressor.compress()` で圧縮

### モジュール責務

| モジュール | 責務 | 型ヒント厳密度 |
| --- | --- | --- |
| `pdf_merge_orchestrator.py` | 全体フロー制御（依存性注入） | 厳密 |
| `document_collector.py` | ディレクトリ走査・ファイル収集・目次構造生成 | 厳密 |
| `pdf_converter.py` | 各種形式→PDF変換のファサード | 厳密 |
| `pdf_processor.py` | PDF操作（マージ、分割、TOC、ブックマーク、ページ番号） | 厳密 |
| `folder_structure_detector.py` | 教育計画（3階層）/行事計画（2階層）の検出 | 厳密 |
| `base_excel_transfer.py` | Excel処理の抽象基底クラス（DRY） | 厳密 |
| `transfer_factory.py` | Excel/Google Sheetsの実装選択（Factory Pattern） | 厳密 |
| `config_loader.py` | 設定読み込み・パス解決・デフォルト値マージ | 通常 |
| `path_validator.py` | パス検証・TOCTOU対策 | 厳密 |
| `exceptions.py` | カスタム例外階層（12種類） | 厳密 |
| `converters/office_converter.py` | Word/Excel/PPT変換（COM） | 緩和（COM型なし） |
| `converters/ichitaro_converter.py` | 一太郎変換（pywinauto） | 緩和（UI自動化） |
| `gui/` | customtkinterベースのGUI | 緩和（tkinter型不完全） |

### 例外階層

```text
PDFMergeError (基底クラス)
├── PDFConversionError
├── ConfigurationError (config_key属性)
├── ResourceError (resource_type属性)
├── FileOperationError (file_path, operation属性)
├── PathNotFoundError (path, description属性)
├── PDFProcessingError (operation属性)
├── ExcelProcessingError (file_path, operation属性)
├── FolderStructureError (directory_path属性)
├── CancelledError
├── GoogleSheetsError (operation属性)
└── GoogleAuthError (auth_stage属性)
```

すべて `original_error` パラメータで例外チェーン可能。

---

## 必須コーディング標準

### 1. 型ヒント

すべての関数・メソッドで型ヒントを使用:

```python
def process_file(file_path: str, max_retries: int = 3) -> Optional[str]:
    """ファイルを処理"""
    pass
```

- `Optional[T]`を明示的に使用
- `Any`は最小限に
- nested関数も戻り値型を明記
- `mypy.ini`で厳密度を制御（コアモジュールは厳密、GUI/COMは緩和）

### 2. 例外処理

統一されたカスタム例外とチェーンを使用:

```python
try:
    process_data()
except ValueError as e:
    raise PDFProcessingError(
        "処理失敗",
        operation="データ処理",
        original_error=e
    ) from e
```

### 3. DRY原則

重複コードは共通メソッドに抽出:

```python
# GUI層ではBaseTabの共通メソッドを使用
canvas, _scrollbar, scrollable_frame = self.create_scrollable_container()

# Excel処理ではBaseExcelTransferを継承
class ExcelTransfer(BaseExcelTransfer):
    def _read_source_data(self) -> List[Dict[str, Any]]:
        # 実装
```

### 4. docstring

Google Styleで記述、Args/Returns/Raisesを明記:

```python
def merge_pdfs(self, pdf_paths: List[str], output_file: str) -> None:
    """
    複数のPDFを1つにマージ

    Args:
        pdf_paths: PDFファイルパスのリスト
        output_file: 出力先ファイルパス

    Raises:
        PDFProcessingError: マージ処理中にエラーが発生した場合
    """
    pass
```

---

## アーキテクチャ原則

### 単一責任原則（SRP）

1つのモジュール/クラスは1つの責務のみ:

- `document_collector.py` → ドキュメント収集
- `pdf_merge_orchestrator.py` → フロー制御
- `pdf_processor.py` → PDF操作

### 依存性注入

コンストラクタで依存を注入し、テスタビリティを確保:

```python
class PDFMergeOrchestrator:
    def __init__(
        self,
        config: ConfigLoader,
        pdf_converter: PDFConverter,
        pdf_processor: PDFProcessor,
        document_collector: DocumentCollector
    ):
        # テスタビリティ向上
```

### デザインパターン

- **Factory Pattern**: `HybridTransferFactory` が Excel/Google Sheets 実装を選択
- **Strategy Pattern**: `BaseExcelTransfer` を継承した複数実装
- **Facade Pattern**: `PDFConverter` が複数の変換機を隠蔽
- **Template Method**: `BaseExcelTransfer` の共通フロー定義

---

## セキュリティ

### PathValidatorの必須使用

ユーザー入力のパスは必ず検証:

```python
is_valid, error_msg, validated_path = PathValidator.validate_directory(
    user_input,
    must_exist=True
)
```

### TOCTOU対策

フラグベースのクリーンアップでTOCTOU脆弱性を回避:

```python
tmp_created = False
try:
    # 処理
    tmp_created = True
finally:
    if tmp_created:
        os.remove(tmp_file)
```

---

## パフォーマンス

### 遅延インポート

**注意**: 頻繁に呼ばれる関数内での遅延インポートは禁止。代わりに`__init__`で依存性注入:

```python
# ✅ Good - コンストラクタで注入
def __init__(self, config, pdf_processor=None):
    self.processor = pdf_processor or PDFProcessor(config)

# ❌ Bad - 毎回インポート
def create_separator_page(self, folder_name: str) -> Optional[str]:
    from pdf_processor import PDFProcessor  # 毎回実行される
    processor = PDFProcessor(self.config)
```

### リソース管理

コンテキストマネージャーを必ず使用:

```python
with fitz.open(pdf_path) as doc:
    # 処理
```

---

## GUI開発

### スクロール可能コンテナ

`BaseTab.create_scrollable_container()`を使用し、重複を排除:

```python
class MyTab(BaseTab):
    def _create_ui(self) -> None:
        canvas, _scrollbar, scrollable_frame = self.create_scrollable_container()
        # scrollable_frameに子ウィジェットを追加
```

### ファイルダイアログ

PathValidatorで必ず検証:

```python
directory = filedialog.askdirectory(title="フォルダを選択")
if directory:
    is_valid, error_msg, validated_path = PathValidator.validate_directory(
        directory, must_exist=True
    )
    if is_valid:
        self.var.set(str(validated_path))
```

---

## Windows固有の考慮事項

### COM初期化（重要）

`run_app.py` でCOMスレッドモデルをSTA（シングルスレッドアパートメント）に設定:

```python
import sys
sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED (STA)
```

これは `tkinter.filedialog` とWin32 COMの両立に必須。

### 一太郎UI自動化

pywinautoベースのUI操作は時間がかかるため、`config.json` でタイミング調整可能:

```json
"ichitaro": {
  "wait_time": 2.0,
  "save_wait": 3.0
}
```

---

## テスト戦略

### pytest マーカー

- `@pytest.mark.unit`: ユニットテスト（モック使用）
- `@pytest.mark.integration`: 統合テスト（実際のファイル操作）
- `@pytest.mark.slow`: 低速テスト（Ghostscript、一太郎など）

### fixtures (`tests/conftest.py`)

- `config`: ConfigLoaderインスタンス
- `temp_pdf`: 一時PDFファイル
- `mock_converter`: PDFConverterモック

### モック推奨箇所

- Win32 COM操作（`office_converter.py`）
- pywinauto UI操作（`ichitaro_converter.py`）
- Google Sheets API（`google_sheets_reader.py`）

---

## 設定管理

### 2種類の設定ファイル

1. **プロジェクト設定** (`config.json`): デフォルト値、ベーステンプレート
2. **ユーザー設定** (`%APPDATA%\PDFMergeSystem\user_config.json`): GUIで編集可能

`ConfigLoader` が両方をマージし、ユーザー設定が優先される。

### 主要設定項目

```json
{
  "year": "2025",                    // 年度（西暦）
  "year_short": "R7",                // 和暦略称（自動計算）
  "transfer_mode": "google_sheets",  // "excel" or "google_sheets"
  "base_paths": {
    "google_drive": "...",
    "local_temp": "..."
  },
  "ghostscript": {
    "executable": "gswin64c.exe"     // 自動検出可能
  }
}
```

---

## Google Sheets統合（v3.5.0）

### OAuth認証フロー

1. `google_auth_manager.py` がトークンライフサイクル管理
2. 初回: ブラウザで認証 → `token.pickle` 保存
3. 2回目以降: トークン自動更新

### Factory による実装切り替え

```python
# transfer_factory.py
factory = HybridTransferFactory(config)
transfer = factory.create_transfer()  # config.transfer_mode で選択

# ExcelTransfer または GoogleSheetsTransfer を返す
```

両方とも `BaseExcelTransfer` を継承しているため、インターフェース統一。

---

## 禁止事項

❌ **過剰エンジニアリング**: 必要のない抽象化層
❌ **後方互換性の破壊**: 既存APIの変更
❌ **グローバル変数**: モジュールレベルの状態
❌ **頻繁関数内の遅延インポート**: パフォーマンス低下

---

## トラブルシューティング

### 開発環境の問題

#### ❌ ModuleNotFoundError: No module named 'XXX'

**原因**: 依存関係が未インストール

**解決策**:

```bash
pip install -r requirements-dev.txt
```

#### ❌ pytest/ruff/mypy: command not found

**原因**: 開発用ツールが未インストール、またはPATHが通っていない

**解決策**:

```bash
# 仮想環境が有効化されているか確認
venv\Scripts\activate

# 開発用依存関係を再インストール
pip install -r requirements-dev.txt
```

#### ❌ pywin32エラー: COM object is not available

**原因**: Microsoft Officeが未インストール、またはCOM登録が不完全

**解決策**:

1. Microsoft Officeがインストールされているか確認
2. 管理者権限でpywin32のCOM登録を実行:

```bash
python venv\Scripts\pywin32_postinstall.py -install
```

### テストの問題

#### ❌ テストが失敗する

**原因**: 環境依存の問題（Office、Ghostscript、一太郎など）

**解決策**:

```bash
# 環境依存テストをスキップ
pytest -m "not slow"

# ユニットテストのみ実行
pytest -m unit
```

### ビルドの問題

#### ❌ PyInstaller: ModuleNotFoundError at runtime

**原因**: `build_installer.spec`の`hiddenimports`に必要なモジュールが未登録

**解決策**:

1. `build_installer.spec`の`hiddenimports`リストに追加
2. クリーンビルド実行: `build.bat`

#### ❌ 実行ファイルサイズが大きい（100MB超）

**原因**: 正常（Python + 依存ライブラリを含むため）

**内訳**:

- Python インタープリタ
- customtkinter、PyPDF2、PyMuPDF、reportlab、Pillow
- pywin32、pywinauto

### Google Sheets連携

#### ❌ 認証情報ファイルが見つかりません

**原因**: `credentials.json`が正しい場所にない

**解決策**:

1. `credentials.json`をプロジェクトルートに配置
2. または`%LOCALAPPDATA%\PDFMergeSystem\credentials.json`に配置
3. 詳細: [docs/GOOGLE_SHEETS_SETUP.md](docs/GOOGLE_SHEETS_SETUP.md)

#### ❌ OAuth認証ブラウザが開かない

**原因**: ファイアウォールでlocalhost通信がブロック

**解決策**:

- IT部門にlocalhost:8080-8090の通信許可を依頼
- または個人ネットワークで初回認証を実施

### 一太郎変換

#### ❌ 一太郎変換が途中で止まる

**原因**: UI自動化のタイミング問題

**解決策**:

`config.json`でタイミングを調整:

```json
{
  "ichitaro": {
    "wait_time": 3.0,
    "save_wait": 5.0
  }
}
```

---

## よくある質問（FAQ）

### FAQ: 開発環境

**Q: macOS/Linuxで開発できますか？**

A: いいえ。Win32 COM、pywinautoを使用するためWindows専用です。

**Q: Python 3.12で動作しますか？**

A: はい。Python 3.8〜3.12で動作確認済みです。

**Q: 仮想環境は必須ですか？**

A: 推奨です。システムPythonとの競合を避けるため、venvまたはcondaの使用を強く推奨します。

### FAQ: テスト

**Q: テストにOfficeインストールは必要ですか？**

A: ユニットテスト（`-m unit`）はモックを使用するため不要です。統合テストは必要です。

**Q: カバレッジ目標は何%ですか？**

A: 明確な目標値はありませんが、コアモジュール（`pdf_merge_orchestrator`、`document_collector`など）は80%以上を推奨します。

### FAQ: ビルド

**Q: インストーラーも作成できますか？**

A: はい。Inno Setupを使用します。詳細は[BUILD_INSTRUCTIONS.md](BUILD_INSTRUCTIONS.md)を参照してください。

**Q: ビルド時間はどのくらいですか？**

A: 環境により異なりますが、通常2〜5分程度です。

### Google Sheets

**Q: 組織アカウント（Google Workspace）で使用できますか？**

A: はい。IT管理者がGoogle Sheets APIを無効にしていない限り使用可能です。

**Q: オフラインで使用できますか？**

A: Google Sheets参照機能はインターネット接続が必要です。Excelモード（`reference_mode: "excel"`）ならオフライン可能です。

### アーキテクチャ

**Q: なぜconvertersディレクトリを分けたのですか？**

A: 単一責任の原則（SRP）に従い、各形式の変換ロジックを独立させるためです（v3.3.0リファクタリング）。

**Q: BaseExcelTransferは何のためですか？**

A: ExcelとGoogle Sheetsの共通ロジックを抽出し、DRY原則を徹底するためです（v3.5.0導入）。

---

## バージョン管理

### コミットメッセージ

```text
種別: 簡潔な説明（50文字以内）

詳細な説明（必要に応じて）

Co-Authored-By: Claude Sonnet 4.5 <noreply@anthropic.com>
```

**種別**: `Feature`, `Fix`, `Refactor`, `Docs`, `Test`, `Chore`, `Build`

### 最近の主要更新

- **v3.5.0** (2026-01): Google Sheets参照機能、BaseExcelTransfer導入（DRY徹底）
- **v3.4.1** (2026-01): Ghostscript自動検出
- **v3.3.0** (2025-12): PDFMergeOrchestrator、セットアップウィザード

---

## クイックリファレンス

### 新規開発者向け：5分でスタート

```bash
# 1. リポジトリをクローン
git clone <repository-url>
cd education-pdf-merger

# 2. 仮想環境セットアップ
python -m venv venv
venv\Scripts\activate

# 3. 依存関係インストール
pip install -r requirements-dev.txt

# 4. アプリケーション起動
python run_app.py

# 5. テスト実行（オプション）
pytest -m unit
```

### よく使うコマンド一覧

| 目的 | コマンド |
| --- | --- |
| アプリ起動 | `python run_app.py` |
| 全テスト実行 | `pytest` |
| ユニットテストのみ | `pytest -m unit` |
| リント・フォーマット | `ruff check --fix . && ruff format .` |
| 型チェック | `mypy --config-file=mypy.ini` |
| EXEビルド | `build.bat` |
| pre-commit手動実行 | `pre-commit run --all-files` |

### 重要なファイル

| ファイル | 役割 |
| --- | --- |
| [run_app.py](run_app.py) | アプリケーションエントリーポイント |
| [pdf_merge_orchestrator.py](pdf_merge_orchestrator.py) | メイン処理フロー制御 |
| [config.json](config.json) | プロジェクト設定（テンプレート） |
| [requirements-dev.txt](requirements-dev.txt) | 開発用依存関係 |
| [mypy.ini](mypy.ini) | 型チェック設定 |
| [pytest.ini](pytest.ini) | テスト設定 |
| [build_installer.spec](build_installer.spec) | PyInstallerビルド設定 |

### 関連ドキュメント

- [README.md](README.md) - プロジェクト概要とユーザー向けガイド
- [CHANGELOG.md](CHANGELOG.md) - バージョン履歴
- [BUILD_INSTRUCTIONS.md](BUILD_INSTRUCTIONS.md) - ビルド手順詳細
- [docs/GOOGLE_SHEETS_SETUP.md](docs/GOOGLE_SHEETS_SETUP.md) - Google Sheets連携セットアップ
- [.github/workflows/ci.yml](.github/workflows/ci.yml) - CI/CD設定

---

このドキュメントは生きたドキュメントです。コーディング方針の改善提案は随時歓迎します。
