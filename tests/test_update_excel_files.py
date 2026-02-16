"""
ExcelTransfer のユニットテスト

Excel転記処理（COM操作をモック化）をテスト
"""
import pytest
from unittest.mock import MagicMock, patch, PropertyMock

from update_excel_files import ExcelTransfer, ExcelTransferError
from exceptions import CancelledError


@pytest.fixture
def transfer():
    """テスト用ExcelTransferインスタンス"""
    return ExcelTransfer(
        ref_filename="C:\\test\\ref.xlsx",
        target_filename="C:\\test\\target.xlsx",
        ref_sheet="Sheet1",
        target_sheet="Sheet1"
    )


@pytest.fixture
def transfer_with_cancel():
    """キャンセル付きExcelTransferインスタンス"""
    return ExcelTransfer(
        ref_filename="ref.xlsx",
        target_filename="target.xlsx",
        ref_sheet="Sheet1",
        target_sheet="Sheet1",
        cancel_check=lambda: True
    )


class TestExcelTransferInit:
    """初期化のテスト"""

    def test_stores_filenames(self, transfer):
        assert transfer.ref_filename == "C:\\test\\ref.xlsx"
        assert transfer.target_filename == "C:\\test\\target.xlsx"

    def test_stores_sheet_names(self, transfer):
        assert transfer.ref_sheet == "Sheet1"
        assert transfer.target_sheet == "Sheet1"

    def test_initial_state(self, transfer):
        assert transfer.excel is None
        assert transfer.ref_wb is None
        assert transfer.target_wb is None
        assert transfer._com_initialized is False

    def test_progress_callback(self):
        cb = MagicMock()
        t = ExcelTransfer("r.xlsx", "t.xlsx", "S1", "S1", progress_callback=cb)
        t._report_progress("test msg")
        cb.assert_called_once_with("test msg")


class TestCheckCancelled:
    """キャンセルチェックのテスト"""

    def test_raises_on_cancel(self, transfer_with_cancel):
        with pytest.raises(CancelledError):
            transfer_with_cancel._check_cancelled()

    def test_no_raise_when_not_cancelled(self, transfer):
        transfer._check_cancelled()  # 例外なし


class TestFindWorkbook:
    """_find_workbook のテスト"""

    @patch('update_excel_files.pythoncom')
    @patch('update_excel_files.win32com.client')
    def test_finds_matching_workbook(self, mock_client, mock_pythoncom, transfer):
        """ファイル名が一致するワークブックを返す"""
        mock_wb = MagicMock()
        mock_wb.Name = "ref.xlsx"

        mock_excel = MagicMock()
        mock_excel.Workbooks = [mock_wb]
        transfer.excel = mock_excel

        result = transfer._find_workbook("C:\\test\\ref.xlsx")
        assert result is mock_wb

    @patch('update_excel_files.pythoncom')
    @patch('update_excel_files.win32com.client')
    def test_raises_when_not_found(self, mock_client, mock_pythoncom, transfer):
        """ワークブックが見つからない場合エラー"""
        mock_excel = MagicMock()
        mock_excel.Workbooks = []
        transfer.excel = mock_excel

        with pytest.raises(ExcelTransferError, match="開かれていません"):
            transfer._find_workbook("missing.xlsx")


class TestConnectWorksheet:
    """_connect_worksheet のテスト"""

    def test_connects_to_sheet(self, transfer):
        """正しいシートに接続できる"""
        mock_wb = MagicMock()
        mock_ws = MagicMock()
        mock_ws_list = [MagicMock()]
        mock_ws_list[0].Name = "Sheet1"
        mock_wb.Worksheets.__iter__ = lambda self: iter(mock_ws_list)
        mock_wb.Worksheets.return_value = mock_ws

        result = transfer._connect_worksheet(mock_wb, "Sheet1", "test.xlsx")
        assert result is mock_ws

    def test_raises_when_sheet_not_found(self, transfer):
        """シートが見つからない場合エラー"""
        mock_wb = MagicMock()
        mock_ws_list = [MagicMock()]
        mock_ws_list[0].Name = "Sheet1"
        mock_wb.Worksheets.__iter__ = lambda self: iter(mock_ws_list)
        mock_wb.Worksheets.side_effect = Exception("Sheet not found")

        with pytest.raises(ExcelTransferError, match="シートが見つかりません"):
            transfer._connect_worksheet(mock_wb, "BadSheet", "test.xlsx")


class TestInitComConnection:
    """_init_com_connection のテスト"""

    @patch('update_excel_files.win32com.client')
    @patch('update_excel_files.pythoncom')
    def test_initializes_com(self, mock_pythoncom, mock_client, transfer):
        """COMが正常に初期化される"""
        mock_excel = MagicMock()
        mock_client.Dispatch.return_value = mock_excel

        transfer._init_com_connection()

        mock_pythoncom.CoInitialize.assert_called_once()
        mock_client.Dispatch.assert_called_once_with("Excel.Application")
        assert transfer.excel is mock_excel
        assert transfer._com_initialized is True

    @patch('update_excel_files.win32com.client')
    @patch('update_excel_files.pythoncom')
    def test_com_already_initialized(self, mock_pythoncom, mock_client, transfer):
        """COM初期化済みの場合はスキップ"""
        mock_pythoncom.CoInitialize.side_effect = Exception("Already initialized")
        mock_client.Dispatch.return_value = MagicMock()

        transfer._init_com_connection()

        assert transfer._com_initialized is False
        assert transfer.excel is not None


class TestCleanupExcel:
    """_cleanup_excel のテスト"""

    @patch('update_excel_files.pythoncom')
    def test_cleanup_releases_all(self, mock_pythoncom, transfer):
        """全COMオブジェクトが解放される"""
        transfer.excel = MagicMock()
        transfer.ref_ws = MagicMock()
        transfer.target_ws = MagicMock()
        transfer.ref_wb = MagicMock()
        transfer.target_wb = MagicMock()
        transfer._com_initialized = True

        transfer._cleanup_excel()

        assert transfer.ref_ws is None
        assert transfer.target_ws is None
        assert transfer.ref_wb is None
        assert transfer.target_wb is None
        assert transfer.excel is None
        assert transfer._com_initialized is False
        mock_pythoncom.CoUninitialize.assert_called_once()

    @patch('update_excel_files.pythoncom')
    def test_cleanup_skips_uninit(self, mock_pythoncom, transfer):
        """COM未初期化時はCoUninitializeをスキップ"""
        transfer._com_initialized = False

        transfer._cleanup_excel()

        mock_pythoncom.CoUninitialize.assert_not_called()


class TestCleanEventNames:
    """_clean_event_names のテスト"""

    def test_cleans_tuple_values(self, transfer):
        """2次元タプルをクリーンなリストに変換"""
        raw = (("入学式",), ("卒業式",), (None,), ("",), ("入学式",))
        result = transfer._clean_event_names(raw)

        assert result == ["入学式", "卒業式"]  # 空・None・重複を除去

    def test_cleans_none_input(self, transfer):
        assert transfer._clean_event_names(None) == []

    def test_cleans_single_value(self, transfer):
        result = transfer._clean_event_names("入学式")
        assert result == ["入学式"]

    def test_strips_whitespace(self, transfer):
        raw = (("  入学式  ",), ("\n卒業式\r",))
        result = transfer._clean_event_names(raw)
        assert result == ["入学式", "卒業式"]


class TestFindValueInSource:
    """_find_value_in_source のテスト"""

    def test_finds_value(self, transfer):
        """検索値が見つかった場合に行番号を返す"""
        mock_ws = MagicMock()
        mock_found = MagicMock()
        mock_found.Row = 15
        mock_ws.Range.return_value.Find.return_value = mock_found
        transfer.ref_ws = mock_ws

        result = transfer._find_value_in_source("入学式")
        assert result == 15

    def test_returns_none_when_not_found(self, transfer):
        """検索値が見つからない場合にNoneを返す"""
        mock_ws = MagicMock()
        mock_ws.Range.return_value.Find.return_value = None
        transfer.ref_ws = mock_ws

        result = transfer._find_value_in_source("存在しない値")
        assert result is None


class TestReadDataRow:
    """_read_data_row のテスト"""

    def test_reads_row_data(self, transfer):
        """行データが正しく読み取られる"""
        mock_ws = MagicMock()
        mock_ws.Range.return_value.Value = (("A", "B", "C"),)
        transfer.ref_ws = mock_ws

        result = transfer._read_data_row(5, "E", "G")
        assert result == ["A", "B", "C"]

    def test_returns_empty_for_none(self, transfer):
        """Noneの場合は空リスト"""
        mock_ws = MagicMock()
        mock_ws.Range.return_value.Value = None
        transfer.ref_ws = mock_ws

        result = transfer._read_data_row(5, "E", "G")
        assert result == []
