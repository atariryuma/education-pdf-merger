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
    ExcelProcessingError
)


class TestPDFMergeError:
    """PDFMergeErrorのテスト"""

    def test_base_exception(self):
        """基底例外クラス"""
        error = PDFMergeError("テストエラー")
        assert str(error) == "テストエラー"
        assert isinstance(error, Exception)


class TestPDFConversionError:
    """PDFConversionErrorのテスト"""

    def test_with_all_params(self):
        """全パラメータ指定"""
        original = ValueError("元のエラー")
        error = PDFConversionError("test.docx", "Word変換", original)
        assert error.file_path == "test.docx"
        assert error.operation == "Word変換"
        assert error.original_error == original
        assert "Word変換失敗" in str(error)
        assert "test.docx" in str(error)

    def test_without_original_error(self):
        """元の例外なし"""
        error = PDFConversionError("test.xlsx", "Excel変換")
        assert error.original_error is None
        assert "Excel変換失敗" in str(error)


class TestConfigurationError:
    """ConfigurationErrorのテスト"""

    def test_with_config_key(self):
        """設定キー指定"""
        error = ConfigurationError("値が無効です", "base_paths.google_drive")
        assert error.config_key == "base_paths.google_drive"
        assert "設定エラー" in str(error)

    def test_without_config_key(self):
        """設定キーなし"""
        error = ConfigurationError("設定ファイルが見つかりません")
        assert error.config_key is None


class TestResourceError:
    """ResourceErrorのテスト"""

    def test_with_all_params(self):
        """全パラメータ指定"""
        original = OSError("接続失敗")
        error = ResourceError("Excel COM", "アプリケーションに接続できません", original)
        assert error.resource_type == "Excel COM"
        assert error.original_error == original
        assert "リソースエラー" in str(error)

    def test_without_original_error(self):
        """元の例外なし"""
        error = ResourceError("ファイル", "ファイルがロックされています")
        assert error.original_error is None


class TestFileOperationError:
    """FileOperationErrorのテスト"""

    def test_read_error(self):
        """読み込みエラー"""
        error = FileOperationError("C:\\test.pdf", "読み込み", IOError("アクセス拒否"))
        assert error.file_path == "C:\\test.pdf"
        assert error.operation == "読み込み"
        assert "ファイル読み込みエラー" in str(error)

    def test_write_error(self):
        """書き込みエラー"""
        error = FileOperationError("C:\\output.pdf", "書き込み")
        assert "ファイル書き込みエラー" in str(error)


class TestPathNotFoundError:
    """PathNotFoundErrorのテスト"""

    def test_with_description(self):
        """説明あり"""
        error = PathNotFoundError("C:\\Documents\\計画書", "教育計画フォルダ")
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
        error = PDFProcessingError("結合", "ページ数が0です")
        assert error.operation == "結合"
        assert "PDF結合エラー" in str(error)

    def test_split_error(self):
        """分割エラー"""
        original = RuntimeError("メモリ不足")
        error = PDFProcessingError("分割", "処理できませんでした", original)
        assert error.original_error == original


class TestExcelProcessingError:
    """ExcelProcessingErrorのテスト"""

    def test_basic_error(self):
        """基本エラー"""
        error = ExcelProcessingError("data.xlsx", "シート削除")
        assert error.file_path == "data.xlsx"
        assert error.operation == "シート削除"
        assert "Excelシート削除エラー" in str(error)

    def test_with_original_error(self):
        """元の例外あり"""
        original = PermissionError("ファイルが使用中")
        error = ExcelProcessingError("locked.xlsx", "更新", original)
        assert error.original_error == original


class TestExceptionHierarchy:
    """例外の継承関係テスト"""

    def test_all_inherit_from_base(self):
        """全例外がPDFMergeErrorを継承"""
        exceptions = [
            PDFConversionError("file", "op"),
            ConfigurationError("msg"),
            ResourceError("type", "msg"),
            FileOperationError("file", "op"),
            PathNotFoundError("path"),
            PDFProcessingError("op", "msg"),
            ExcelProcessingError("file", "op"),
        ]
        for exc in exceptions:
            assert isinstance(exc, PDFMergeError)
            assert isinstance(exc, Exception)

    def test_can_be_caught_by_base(self):
        """基底クラスでキャッチ可能"""
        try:
            raise PDFConversionError("test.docx", "変換")
        except PDFMergeError as e:
            assert True
        except Exception:
            pytest.fail("PDFMergeErrorでキャッチされるべき")
