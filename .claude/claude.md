# 教育計画PDFマージシステム - コーディング方針

> **最終更新**: 2026-01-16
> **バージョン**: 3.4.0

## プロジェクト概要

教育計画や行事計画のドキュメントを自動的にPDF化して統合するデスクトップアプリケーション。
Word、Excel、PowerPoint、一太郎、画像ファイルなどをPDFに変換し、目次付きの単一PDFファイルを生成します。

## コーディング標準

### 1. 型ヒント（Type Hints）

**必須**: すべての関数・メソッドで型ヒントを使用

```python
# ✅ Good
def process_file(file_path: str, max_retries: int = 3) -> Optional[str]:
    """ファイルを処理"""
    pass

# ❌ Bad
def process_file(file_path, max_retries=3):
    """ファイルを処理"""
    pass
```

**型ヒントのベストプラクティス**:
- `Optional[T]`を明示的に使用（`None`の可能性がある場合）
- `Union`より`Optional`を優先
- `Any`は最小限に（やむを得ない場合のみ）
- `TYPE_CHECKING`を使って循環インポートを回避

### 2. 例外処理

**統一されたカスタム例外**を使用:

```python
from exceptions import PDFProcessingError, ConfigurationError

# ✅ Good - キーワード専用引数を使用
raise PDFProcessingError(
    "PDFの読み込みに失敗しました",
    operation="読み込み",
    original_error=e
) from e

# ❌ Bad - 位置引数（旧API）
raise PDFProcessingError("読み込み", "PDFの読み込みに失敗しました", e)
```

**例外チェーン**を必ず使用:
```python
try:
    process_data()
except ValueError as e:
    raise CustomError("処理失敗", original_error=e) from e
```

### 3. 定数管理

**マジックナンバー禁止** - すべて定数化:

```python
from constants import PDFConstants

# ✅ Good
current_page = PDFConstants.CONTENT_START_PAGE

# ❌ Bad
current_page = 3  # 表紙 + 目次 + 1
```

**定数の配置**:
- `constants.py`にグローバル定数を集約
- クラス固有の定数はクラス内に定義しない（定数クラスを使用）

### 4. DRY原則（Don't Repeat Yourself）

**コンテキストマネージャーでパターンを抽象化**:

```python
# ✅ Good - 重複排除
with self._atomic_pdf_operation(pdf_path) as tmp_file:
    # 処理
    doc.save(tmp_file)

# ❌ Bad - 重複コード
tmp_file = pdf_path + ".tmp"
try:
    # 処理
    doc.save(tmp_file)
    os.replace(tmp_file, pdf_path)
finally:
    if os.path.exists(tmp_file):
        os.remove(tmp_file)
```

### 5. モジュール構造

**単一責任原則（SRP）**を徹底:

```
✅ Good - 責務が分離されている
- document_collector.py     # ドキュメント収集のみ
- pdf_merge_orchestrator.py # 全体のフロー制御のみ
- pdf_processor.py           # PDF操作のみ

❌ Bad - 複数の責務が混在
- document_collector.py      # 収集 + オーケストレーション
```

### 6. docstring

**Google スタイル**を使用:

```python
def merge_pdfs(self, pdf_paths: List[str], output_file: str) -> None:
    """
    複数のPDFを1つにマージ

    Args:
        pdf_paths: PDFファイルパスのリスト
        output_file: 出力先ファイルパス

    Raises:
        PDFProcessingError: マージ処理中にエラーが発生した場合

    Note:
        ファイルが存在しない場合はスキップされます
    """
    pass
```

### 7. ロギング

**構造化ロギング**を推奨:

```python
# ✅ Good - コンテキスト情報を含む
logger.info(f"PDF統合完了: {output_path} (ページ数: {total_pages})")

# ⚠️ Acceptable - 最小限の情報
logger.info("PDF統合完了")

# ❌ Bad - デバッグ情報不足
logger.info("完了")
```

**ログレベルの使い分け**:
- `DEBUG`: 開発時のトレース情報
- `INFO`: 通常の処理フロー
- `WARNING`: 異常だが続行可能
- `ERROR`: エラーが発生（処理継続不可）
- `CRITICAL`: システム全体に影響

### 8. パス処理

**pathlib優先**（ただし既存コードとの互換性を維持）:

```python
from pathlib import Path

# ✅ Good
path = Path(file_path)
if path.exists() and path.is_file():
    process(str(path))

# ⚠️ Acceptable - 既存コードとの互換性
if os.path.exists(file_path) and os.path.isfile(file_path):
    process(file_path)
```

## アーキテクチャパターン

### 1. ファサードパターン

`PDFConverter`は各種変換器のファサード:

```python
class PDFConverter:
    def __init__(self, temp_dir, ...):
        self.office_converter = OfficeConverter(temp_dir)
        self.image_converter = ImageConverter()
        self.ichitaro_converter = IchitaroConverter(...)

    def convert(self, file_path: str, output_path: Optional[str] = None) -> Optional[str]:
        ext = os.path.splitext(file_path)[1].lower()
        if ext in self.OFFICE_EXTENSIONS:
            return self.office_converter.convert(file_path, output_path)
        # ...
```

### 2. テンプレートメソッドパターン

共通パターンの抽象化:

```python
@contextmanager
def _atomic_pdf_operation(self, pdf_path: str) -> Generator[str, None, None]:
    """一時ファイルを使った安全なPDF操作"""
    tmp_file = pdf_path + PDFConstants.TEMP_FILE_SUFFIX
    try:
        yield tmp_file
        os.replace(tmp_file, pdf_path)
    finally:
        # クリーンアップ
```

### 3. 依存性注入

コンストラクタで依存を注入:

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

## セキュリティ

### 1. パス検証

**PathValidator**を必ず使用:

```python
from path_validator import PathValidator

# ✅ Good
is_valid, error_msg, validated_path = PathValidator.validate_directory(
    user_input,
    must_exist=True
)

# ❌ Bad - 検証なし
if os.path.exists(user_input):
    process(user_input)
```

### 2. ファイル名サニタイズ

**無効文字の除去**:

```python
safe_name = PathValidator.sanitize_filename(
    user_input,
    replacement='_',
    default_name='default'
)
```

### 3. TOCTOU対策

**フラグベースのクリーンアップ**:

```python
# ✅ Good - フラグでクリーンアップ要否を判定
tmp_created = False
try:
    # 処理
    tmp_created = True
finally:
    if tmp_created:
        os.remove(tmp_file)

# ❌ Bad - TOCTOU脆弱性
finally:
    if os.path.exists(tmp_file):
        os.remove(tmp_file)
```

## パフォーマンス

### 1. 遅延インポート

**循環インポート回避**と**起動時間短縮**:

```python
def create_separator_page(self, folder_name: str) -> Optional[str]:
    # 必要なときだけインポート
    from pdf_processor import PDFProcessor
    processor = PDFProcessor(self.config)
    return processor.create_separator_pdf(folder_name, output_pdf)
```

### 2. リソース管理

**コンテキストマネージャー必須**:

```python
# ✅ Good
with fitz.open(pdf_path) as doc:
    # 処理

# ❌ Bad
doc = fitz.open(pdf_path)
# 処理
doc.close()  # 例外時にリークの可能性
```

## テスト（将来対応）

### 優先度高

1. **ユニットテスト**: 各モジュールの単体テスト
2. **統合テスト**: PDF生成の E2E テスト
3. **例外処理テスト**: エラーケースの網羅

```python
# tests/test_pdf_processor.py (将来)
def test_merge_pdfs():
    processor = PDFProcessor(config)
    result = processor.merge_pdfs(pdf_list, output)
    assert os.path.exists(output)
```

## 禁止事項

❌ **過剰エンジニアリング禁止**:
- 必要のない抽象化層
- 使われない拡張ポイント
- 複雑すぎるデザインパターン

❌ **後方互換性を破壊する変更**:
- 既存APIの引数順序変更
- publicメソッドの削除
- 既存機能の削除

❌ **グローバル変数**:
- モジュールレベルの状態
- シングルトンの乱用

## バージョン管理

### コミットメッセージ

```
種別: 簡潔な説明（50文字以内）

詳細な説明（必要に応じて）

Co-Authored-By: Claude Sonnet 4.5 <noreply@anthropic.com>
```

**種別**:
- `Feature`: 新機能
- `Fix`: バグ修正
- `Refactor`: リファクタリング
- `Docs`: ドキュメント更新
- `Test`: テスト追加・修正
- `Chore`: ビルド・設定変更

### ブランチ戦略

- `main`: 本番リリース用
- 直接コミット可（小規模プロジェクト）

## 参考リソース

- **Python公式スタイルガイド**: PEP 8
- **型ヒント**: PEP 484, PEP 526
- **docstring**: Google Style Guide
- **セキュリティ**: OWASP Top 10

---

**このドキュメントは生きたドキュメントです**
コーディング方針の改善提案は随時歓迎します。
