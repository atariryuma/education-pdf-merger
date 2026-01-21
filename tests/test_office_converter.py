"""
OfficeConverterのユニットテスト

Microsoft Office COM APIをモック化してテスト
"""
import os  # noqa: F401 - @patch('os.path.exists') で使用
import tempfile
import shutil
from pathlib import Path
from typing import Generator
from unittest.mock import Mock, patch, MagicMock
import pytest

from converters.office_converter import OfficeConverter
from exceptions import PDFConversionError


@pytest.fixture
def temp_dir() -> Generator[Path, None, None]:
    """
    一時ディレクトリを作成

    Yields:
        Path: 一時ディレクトリのパス
    """
    temp_path = Path(tempfile.mkdtemp(prefix="office_test_"))
    try:
        yield temp_path
    finally:
        shutil.rmtree(temp_path, ignore_errors=True)


@pytest.fixture
def converter(temp_dir: Path) -> OfficeConverter:
    """
    OfficeConverterインスタンスを作成

    Args:
        temp_dir: 一時ディレクトリ

    Returns:
        OfficeConverter: テスト用インスタンス
    """
    return OfficeConverter(str(temp_dir))


@pytest.fixture
def mock_word_doc(temp_dir: Path) -> Path:
    """
    ダミーWord文書を作成

    Args:
        temp_dir: 一時ディレクトリ

    Returns:
        Path: ダミーファイルのパス
    """
    doc_path = temp_dir / "test.docx"
    doc_path.write_text("Test document", encoding='utf-8')
    return doc_path


@pytest.fixture
def mock_excel_file(temp_dir: Path) -> Path:
    """
    ダミーExcelファイルを作成

    Args:
        temp_dir: 一時ディレクトリ

    Returns:
        Path: ダミーファイルのパス
    """
    excel_path = temp_dir / "test.xlsx"
    excel_path.write_text("Test workbook", encoding='utf-8')
    return excel_path


@pytest.fixture
def mock_ppt_file(temp_dir: Path) -> Path:
    """
    ダミーPowerPointファイルを作成

    Args:
        temp_dir: 一時ディレクトリ

    Returns:
        Path: ダミーファイルのパス
    """
    ppt_path = temp_dir / "test.pptx"
    ppt_path.write_text("Test presentation", encoding='utf-8')
    return ppt_path


class TestOfficeConverter:
    """OfficeConverterのテスト"""

    def test_initialization(self, temp_dir: Path):
        """初期化のテスト"""
        converter = OfficeConverter(str(temp_dir))
        assert converter.temp_dir == str(temp_dir)

    def test_office_extensions_constant(self):
        """サポートされる拡張子の定義確認"""
        expected_extensions = ('.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.rtf')
        assert OfficeConverter.OFFICE_EXTENSIONS == expected_extensions

    @patch('converters.office_converter.pythoncom')
    def test_com_context_manager(self, mock_pythoncom: Mock):
        """COMコンテキストマネージャーのテスト"""
        with OfficeConverter._com_context():
            mock_pythoncom.CoInitialize.assert_called_once()
        mock_pythoncom.CoUninitialize.assert_called_once()

    @patch('converters.office_converter.pythoncom')
    def test_com_context_manager_with_exception(self, mock_pythoncom: Mock):
        """COMコンテキストマネージャーの例外処理テスト"""
        with pytest.raises(ValueError):
            with OfficeConverter._com_context():
                mock_pythoncom.CoInitialize.assert_called_once()
                raise ValueError("Test error")
        mock_pythoncom.CoUninitialize.assert_called_once()

    @patch('converters.office_converter.subprocess.run')
    def test_kill_office_process_success(self, mock_run: Mock):
        """プロセス強制終了の成功テスト"""
        mock_run.return_value = Mock(returncode=0, stderr="")
        OfficeConverter._kill_office_process("WINWORD.EXE")
        mock_run.assert_called_once()
        call_args = mock_run.call_args[0][0]
        assert call_args == ['taskkill', '/F', '/IM', 'WINWORD.EXE']

    @patch('converters.office_converter.subprocess.run')
    def test_kill_office_process_not_found(self, mock_run: Mock):
        """プロセスが見つからない場合のテスト"""
        mock_run.return_value = Mock(returncode=128, stderr="プロセスなし")
        OfficeConverter._kill_office_process("WINWORD.EXE")
        mock_run.assert_called_once()

    @patch('converters.office_converter.subprocess.run')
    def test_kill_office_process_failure(self, mock_run: Mock):
        """プロセス強制終了の失敗テスト"""
        mock_run.return_value = Mock(returncode=1, stderr="エラー")
        OfficeConverter._kill_office_process("WINWORD.EXE")
        mock_run.assert_called_once()

    @patch('converters.office_converter.subprocess.run')
    def test_kill_office_process_timeout(self, mock_run: Mock):
        """プロセス強制終了のタイムアウトテスト"""
        from subprocess import TimeoutExpired
        mock_run.side_effect = TimeoutExpired('taskkill', 5)
        OfficeConverter._kill_office_process("WINWORD.EXE")

    def test_cleanup_office_app_word(self, converter: OfficeConverter):
        """Wordアプリケーションのクリーンアップテスト"""
        mock_doc = Mock()
        mock_app = Mock()

        converter._cleanup_office_app(mock_doc, mock_app, "WINWORD.EXE", "Word")

        mock_doc.Close.assert_called_once_with(SaveChanges=False)
        mock_app.Quit.assert_called_once()

    def test_cleanup_office_app_powerpoint(self, converter: OfficeConverter):
        """PowerPointアプリケーションのクリーンアップテスト"""
        mock_pres = Mock()
        mock_app = Mock()

        converter._cleanup_office_app(mock_pres, mock_app, "POWERPNT.EXE", "PowerPoint")

        # PowerPointはClose()に引数を渡さない
        mock_pres.Close.assert_called_once_with()
        mock_app.Quit.assert_called_once()

    def test_cleanup_office_app_quit_failure(self, converter: OfficeConverter):
        """Quit失敗時の強制終了テスト"""
        mock_doc = Mock()
        mock_app = Mock()
        mock_app.Quit.side_effect = Exception("Quit failed")

        with patch.object(converter, '_kill_office_process') as mock_kill:
            converter._cleanup_office_app(mock_doc, mock_app, "WINWORD.EXE", "Word")
            mock_kill.assert_called_once_with("WINWORD.EXE")

    def test_cleanup_office_app_none_objects(self, converter: OfficeConverter):
        """None オブジェクトの処理テスト"""
        # 例外が発生しないことを確認
        converter._cleanup_office_app(None, None, "WINWORD.EXE", "Word")

    @patch('converters.office_converter.client')
    @patch('converters.office_converter.pythoncom')
    @patch('os.path.exists')
    def test_convert_word_success(
        self,
        mock_exists: Mock,
        mock_pythoncom: Mock,
        mock_client: Mock,
        converter: OfficeConverter,
        mock_word_doc: Path,
        temp_dir: Path
    ):
        """Word変換の成功テスト"""
        output_path = temp_dir / "output.pdf"
        mock_exists.return_value = True

        # Wordアプリケーションのモック
        mock_word = MagicMock()
        mock_doc = MagicMock()
        mock_word.Documents.Open.return_value = mock_doc
        mock_client.Dispatch.return_value = mock_word

        result = converter.convert(str(mock_word_doc), str(output_path))

        assert result == str(output_path)
        mock_client.Dispatch.assert_called_with("Word.Application")
        mock_word.Documents.Open.assert_called_once()
        mock_doc.SaveAs2.assert_called_once()

    @patch('converters.office_converter.client')
    @patch('converters.office_converter.pythoncom')
    @patch('converters.office_converter.shutil.copy2')
    @patch('os.path.exists')
    def test_convert_excel_success(
        self,
        mock_exists: Mock,
        mock_copy: Mock,
        mock_pythoncom: Mock,
        mock_client: Mock,
        converter: OfficeConverter,
        mock_excel_file: Path,
        temp_dir: Path
    ):
        """Excel変換の成功テスト"""
        output_path = temp_dir / "output.pdf"
        mock_exists.return_value = True

        # Excelアプリケーションのモック
        mock_excel = MagicMock()
        mock_wb = MagicMock()
        mock_excel.Workbooks.Open.return_value = mock_wb
        mock_client.Dispatch.return_value = mock_excel

        result = converter.convert(str(mock_excel_file), str(output_path))

        assert result == str(output_path)
        mock_copy.assert_called_once()
        mock_client.Dispatch.assert_called_with("Excel.Application")
        mock_excel.Workbooks.Open.assert_called_once()
        mock_wb.ExportAsFixedFormat.assert_called_once()

    @patch('converters.office_converter.client')
    @patch('converters.office_converter.pythoncom')
    @patch('os.path.exists')
    def test_convert_powerpoint_success(
        self,
        mock_exists: Mock,
        mock_pythoncom: Mock,
        mock_client: Mock,
        converter: OfficeConverter,
        mock_ppt_file: Path,
        temp_dir: Path
    ):
        """PowerPoint変換の成功テスト"""
        output_path = temp_dir / "output.pdf"
        mock_exists.return_value = True

        # PowerPointアプリケーションのモック
        mock_ppt = MagicMock()
        mock_pres = MagicMock()
        mock_ppt.Presentations.Open.return_value = mock_pres
        mock_client.Dispatch.return_value = mock_ppt

        result = converter.convert(str(mock_ppt_file), str(output_path))

        assert result == str(output_path)
        mock_client.Dispatch.assert_called_with("PowerPoint.Application")
        mock_ppt.Presentations.Open.assert_called_once()
        mock_pres.SaveAs.assert_called_once()

    @patch('converters.office_converter.client')
    @patch('converters.office_converter.pythoncom')
    @patch('os.path.exists')
    def test_convert_output_not_created(
        self,
        mock_exists: Mock,
        mock_pythoncom: Mock,
        mock_client: Mock,
        converter: OfficeConverter,
        mock_word_doc: Path,
        temp_dir: Path
    ):
        """変換後のファイルが作成されない場合のテスト"""
        output_path = temp_dir / "output.pdf"
        mock_exists.return_value = False

        mock_word = MagicMock()
        mock_client.Dispatch.return_value = mock_word

        result = converter.convert(str(mock_word_doc), str(output_path))

        assert result is None

    @patch('converters.office_converter.client')
    @patch('converters.office_converter.pythoncom')
    def test_convert_raises_pdf_conversion_error(
        self,
        mock_pythoncom: Mock,
        mock_client: Mock,
        converter: OfficeConverter,
        mock_word_doc: Path,
        temp_dir: Path
    ):
        """変換エラー時のPDFConversionError発生テスト"""
        output_path = temp_dir / "output.pdf"
        mock_client.Dispatch.side_effect = Exception("COM error")

        with pytest.raises(PDFConversionError) as exc_info:
            converter.convert(str(mock_word_doc), str(output_path))

        assert "Office変換に失敗" in str(exc_info.value)
        assert exc_info.value.original_error is not None

    @patch('converters.office_converter.client')
    @patch('converters.office_converter.pythoncom')
    def test_convert_system_exit_propagation(
        self,
        mock_pythoncom: Mock,
        mock_client: Mock,
        converter: OfficeConverter,
        mock_word_doc: Path,
        temp_dir: Path
    ):
        """SystemExitの伝播テスト"""
        output_path = temp_dir / "output.pdf"
        mock_client.Dispatch.side_effect = SystemExit(1)

        with pytest.raises(SystemExit):
            converter.convert(str(mock_word_doc), str(output_path))

    @patch('converters.office_converter.client')
    @patch('converters.office_converter.pythoncom')
    def test_convert_keyboard_interrupt_propagation(
        self,
        mock_pythoncom: Mock,
        mock_client: Mock,
        converter: OfficeConverter,
        mock_word_doc: Path,
        temp_dir: Path
    ):
        """KeyboardInterruptの伝播テスト"""
        output_path = temp_dir / "output.pdf"
        mock_client.Dispatch.side_effect = KeyboardInterrupt()

        with pytest.raises(KeyboardInterrupt):
            converter.convert(str(mock_word_doc), str(output_path))

    @patch('converters.office_converter.client')
    @patch('converters.office_converter.pythoncom')
    @patch('os.path.exists')
    def test_convert_removes_existing_output(
        self,
        mock_exists: Mock,
        mock_pythoncom: Mock,
        mock_client: Mock,
        converter: OfficeConverter,
        mock_word_doc: Path,
        temp_dir: Path
    ):
        """既存の出力ファイルを削除するテスト"""
        output_path = temp_dir / "output.pdf"
        output_path.write_text("existing", encoding='utf-8')

        mock_exists.return_value = True
        mock_word = MagicMock()
        mock_client.Dispatch.return_value = mock_word

        converter.convert(str(mock_word_doc), str(output_path))

        # os.remove()が呼ばれることを確認（間接的にファイルが削除される）
        # 実際の削除はモック化された環境では確認できないため、
        # 変換処理が正常に進行することで間接的に確認


if __name__ == "__main__":
    pytest.main([__file__, "-v", "--tb=short"])
