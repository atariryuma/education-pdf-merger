# 教育計画PDFマージシステム - コーディング方針

> **最終更新**: 2026-01-21
> **バージョン**: 3.4.1

## プロジェクト概要

教育計画や行事計画のドキュメントをPDF化して統合するデスクトップアプリケーション。
Word、Excel、PowerPoint、一太郎、画像ファイルを目次付き単一PDFに変換します。

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

## 禁止事項

❌ **過剰エンジニアリング**: 必要のない抽象化層
❌ **後方互換性の破壊**: 既存APIの変更
❌ **グローバル変数**: モジュールレベルの状態

---

## バージョン管理

### コミットメッセージ

```
種別: 簡潔な説明（50文字以内）

詳細な説明（必要に応じて）

Co-Authored-By: Claude Sonnet 4.5 <noreply@anthropic.com>
```

**種別**: `Feature`, `Fix`, `Refactor`, `Docs`, `Test`, `Chore`

---

このドキュメントは生きたドキュメントです。コーディング方針の改善提案は随時歓迎します。
