"""
カスタム例外クラスのテスト
"""
import os
import pytest
import sys

# プロジェクトルートをパスに追加
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from exceptions import (
    PDFMergeError,
    PDFConversionError,
    ConfigurationError,
    ResourceError,
    FileOperationError,
    PathNotFoundError,
    PDFProcessingError,
    ExcelProcessingError,
    CancelledError
)


class TestPDFMergeError:
    """PDFMergeErrorのテスト"""

    def test_base_exception(self):
        """基底例外クラス"""
        error = PDFMergeError("テストエラー")
        assert str(error) == "テストエラー"
        assert isinstance(error, Exception)

    def test_with_original_error(self):
        """元の例外あり"""
        original = ValueError("元のエラー")
        error = PDFMergeError("テストエラー", original_error=original)
        assert error.original_error == original
        assert "ValueError" in str(error)


class TestPDFConversionError:
    """PDFConversionErrorのテスト"""

    def test_with_all_params(self):
        """全パラメータ指定"""
        original = ValueError("元のエラー")
        error = PDFConversionError(
            "Word変換失敗: test.docx",
            original_error=original,
            file_path="test.docx",
            operation="Word変換"
        )
        assert error.file_path == "test.docx"
        assert error.operation == "Word変換"
        assert error.original_error == original
        assert "Word変換失敗" in str(error)
        assert "test.docx" in str(error)

    def test_without_original_error(self):
        """元の例外なし"""
        error = PDFConversionError(
            "Excel変換失敗: test.xlsx",
            file_path="test.xlsx",
            operation="Excel変換"
        )
        assert error.original_error is None
        assert "Excel変換失敗" in str(error)


class TestConfigurationError:
    """ConfigurationErrorのテスト"""

    def test_with_config_key(self):
        """設定キー指定"""
        error = ConfigurationError("値が無効です", config_key="base_paths.google_drive")
        assert error.config_key == "base_paths.google_drive"
        assert "設定エラー" in str(error)

    def test_without_config_key(self):
        """設定キーなし"""
        error = ConfigurationError("設定ファイルが見つかりません")
        assert not hasattr(error, 'config_key') or error.config_key is None


class TestResourceError:
    """ResourceErrorのテスト"""

    def test_with_all_params(self):
        """全パラメータ指定"""
        original = OSError("接続失敗")
        error = ResourceError(
            "アプリケーションに接続できません",
            resource_type="Excel COM",
            original_error=original
        )
        assert error.resource_type == "Excel COM"
        assert error.original_error == original
        assert "リソースエラー" in str(error)

    def test_without_original_error(self):
        """元の例外なし"""
        error = ResourceError(
            "ファイルがロックされています",
            resource_type="ファイル"
        )
        assert error.original_error is None


class TestFileOperationError:
    """FileOperationErrorのテスト"""

    def test_read_error(self):
        """読み込みエラー"""
        original = IOError("アクセス拒否")
        error = FileOperationError(
            "アクセス拒否",
            file_path="C:\\test.pdf",
            operation="読み込み",
            original_error=original
        )
        assert error.file_path == "C:\\test.pdf"
        assert error.operation == "読み込み"
        assert "ファイル読み込みエラー" in str(error)

    def test_write_error(self):
        """書き込みエラー"""
        error = FileOperationError(
            "ディスク容量不足",
            file_path="C:\\output.pdf",
            operation="書き込み"
        )
        assert "ファイル書き込みエラー" in str(error)


class TestPathNotFoundError:
    """PathNotFoundErrorのテスト"""

    def test_with_description(self):
        """説明あり"""
        error = PathNotFoundError(
            "C:\\Documents\\計画書",
            description="教育計画フォルダ"
        )
        assert error.path == "C:\\Documents\\計画書"
        assert error.description == "教育計画フォルダ"
        assert "教育計画フォルダが見つかりません" in str(error)

    def test_without_description(self):
        """説明なし"""
        error = PathNotFoundError("C:\\NonExistent")
        assert error.description is None
        assert "パスが見つかりません" in str(error)


class TestPDFProcessingError:
    """PDFProcessingErrorのテスト"""

    def test_merge_error(self):
        """結合エラー"""
        error = PDFProcessingError("ページ数が0です", operation="結合")
        assert error.operation == "結合"
        assert "PDF結合エラー" in str(error)

    def test_split_error(self):
        """分割エラー"""
        original = RuntimeError("メモリ不足")
        error = PDFProcessingError(
            "処理できませんでした",
            operation="分割",
            original_error=original
        )
        assert error.original_error == original


class TestExcelProcessingError:
    """ExcelProcessingErrorのテスト"""

    def test_basic_error(self):
        """基本エラー"""
        error = ExcelProcessingError(
            "シートが見つかりません",
            file_path="data.xlsx",
            operation="シート削除"
        )
        assert error.file_path == "data.xlsx"
        assert error.operation == "シート削除"
        assert "Excelシート削除エラー" in str(error)

    def test_with_original_error(self):
        """元の例外あり"""
        original = PermissionError("ファイルが使用中")
        error = ExcelProcessingError(
            "ファイルが使用中です",
            file_path="locked.xlsx",
            operation="更新",
            original_error=original
        )
        assert error.original_error == original


class TestCancelledError:
    """CancelledErrorのテスト"""

    def test_default_message(self):
        """デフォルトメッセージ"""
        error = CancelledError()
        assert "キャンセル" in str(error)

    def test_custom_message(self):
        """カスタムメッセージ"""
        error = CancelledError("一太郎変換がキャンセルされました")
        assert "一太郎変換がキャンセルされました" in str(error)


class TestExceptionHierarchy:
    """例外の継承関係テスト"""

    def test_all_inherit_from_base(self):
        """全例外がPDFMergeErrorを継承"""
        exceptions = [
            PDFConversionError("変換失敗"),
            ConfigurationError("設定エラー"),
            ResourceError("リソースエラー", resource_type="ファイル"),
            FileOperationError("操作エラー", file_path="file.pdf", operation="読み込み"),
            PathNotFoundError("path"),
            PDFProcessingError("処理エラー", operation="結合"),
            ExcelProcessingError("Excelエラー", file_path="file.xlsx", operation="更新"),
        ]
        for exc in exceptions:
            assert isinstance(exc, PDFMergeError)
            assert isinstance(exc, Exception)

    def test_can_be_caught_by_base(self):
        """基底クラスでキャッチ可能"""
        try:
            raise PDFConversionError("変換エラー")
        except PDFMergeError:
            assert True
        except Exception:
            pytest.fail("PDFMergeErrorでキャッチされるべき")
