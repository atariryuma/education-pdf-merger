"""infrastructure.com_utils のユニットテスト"""
from unittest.mock import patch, MagicMock

import pytest

from infrastructure.com_utils import com_apartment


@pytest.mark.unit
class TestComApartment:
    """com_apartment コンテキストマネージャのテスト"""

    @patch("infrastructure.com_utils.pythoncom")
    def test_sta_calls_co_initialize_ex(self, mock_pythoncom: MagicMock):
        """sta=True で CoInitializeEx(COINIT_APARTMENTTHREADED) を呼ぶ"""
        mock_pythoncom.COINIT_APARTMENTTHREADED = 2
        with com_apartment(sta=True):
            mock_pythoncom.CoInitializeEx.assert_called_once_with(2)
            mock_pythoncom.CoInitialize.assert_not_called()
        mock_pythoncom.CoUninitialize.assert_called_once()

    @patch("infrastructure.com_utils.pythoncom")
    def test_mta_calls_co_initialize(self, mock_pythoncom: MagicMock):
        """sta=False で CoInitialize を呼ぶ"""
        with com_apartment(sta=False):
            mock_pythoncom.CoInitialize.assert_called_once()
            mock_pythoncom.CoInitializeEx.assert_not_called()
        mock_pythoncom.CoUninitialize.assert_called_once()

    @patch("infrastructure.com_utils.pythoncom")
    def test_uninitialize_called_on_exception(self, mock_pythoncom: MagicMock):
        """例外発生時もCoUninitializeが呼ばれる"""
        with pytest.raises(RuntimeError):
            with com_apartment(sta=False):
                raise RuntimeError("test")
        mock_pythoncom.CoUninitialize.assert_called_once()

    @patch("infrastructure.com_utils.pythoncom")
    def test_uninitialize_exception_is_swallowed(self, mock_pythoncom: MagicMock):
        """CoUninitializeで例外が発生しても呼び元に伝播しない（多重解放対策）"""
        mock_pythoncom.CoUninitialize.side_effect = OSError("already uninitialized")
        # 例外なく抜けられること
        with com_apartment(sta=False):
            pass
