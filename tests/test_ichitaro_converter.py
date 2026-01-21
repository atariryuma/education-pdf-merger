"""
IchitaroConverterのユニットテスト

pywinautoをモック化してテスト
"""
import os  # noqa: F401 - @patch('os.path.exists'), @patch('os.startfile') で使用
import tempfile
import shutil
import time
from pathlib import Path
from typing import Generator
from unittest.mock import Mock, patch, MagicMock
import pytest

from converters.ichitaro_converter import IchitaroConverter
from exceptions import CancelledError


@pytest.fixture
def temp_dir() -> Generator[Path, None, None]:
    """
    一時ディレクトリを作成

    Yields:
        Path: 一時ディレクトリのパス
    """
    temp_path = Path(tempfile.mkdtemp(prefix="ichitaro_test_"))
    try:
        yield temp_path
    finally:
        shutil.rmtree(temp_path, ignore_errors=True)


@pytest.fixture
def converter() -> IchitaroConverter:
    """
    IchitaroConverterインスタンスを作成

    Returns:
        IchitaroConverter: テスト用インスタンス
    """
    return IchitaroConverter()


@pytest.fixture
def converter_with_cancel() -> IchitaroConverter:
    """
    キャンセルチェック付きIchitaroConverterを作成

    Returns:
        IchitaroConverter: キャンセル可能なインスタンス
    """
    cancel_flag = False

    def cancel_check():
        return cancel_flag

    return IchitaroConverter(cancel_check=cancel_check)


@pytest.fixture
def mock_jtd_file(temp_dir: Path) -> Path:
    """
    ダミー一太郎ファイルを作成

    Args:
        temp_dir: 一時ディレクトリ

    Returns:
        Path: ダミーファイルのパス
    """
    jtd_path = temp_dir / "test.jtd"
    jtd_path.write_text("一太郎ダミー文書", encoding='utf-8')
    return jtd_path


class TestIchitaroConverter:
    """IchitaroConverterのテスト"""

    def test_initialization_default(self):
        """デフォルト初期化のテスト"""
        converter = IchitaroConverter()
        assert converter.ichitaro_settings is not None
        assert converter.ichitaro_settings['printer_name'] == 'Microsoft Print to PDF'
        assert not converter.is_cancelled()

    def test_initialization_with_settings(self):
        """カスタム設定での初期化テスト"""
        custom_settings = {
            'ichitaro_ready_timeout': 60,
            'save_wait_seconds': 30,
            'printer_name': 'Custom PDF Printer'
        }
        converter = IchitaroConverter(ichitaro_settings=custom_settings)
        assert converter.ichitaro_settings['ichitaro_ready_timeout'] == 60
        assert converter.ichitaro_settings['save_wait_seconds'] == 30

    def test_initialization_with_cancel_check(self):
        """キャンセルチェック付き初期化テスト"""
        cancel_state = {'cancelled': False}

        def cancel_check():
            return cancel_state['cancelled']

        converter = IchitaroConverter(cancel_check=cancel_check)
        assert not converter.is_cancelled()

        cancel_state['cancelled'] = True
        assert converter.is_cancelled()  # cancel_stateは辞書なので変更が反映される

    def test_initialization_with_dialog_callback(self):
        """ダイアログコールバック付き初期化テスト"""
        callback_called = []

        def dialog_callback(message: str, show: bool):
            callback_called.append((message, show))

        converter = IchitaroConverter(dialog_callback=dialog_callback)
        assert converter._dialog_callback is dialog_callback

    def test_ichitaro_extensions_constant(self):
        """サポートされる拡張子の定義確認"""
        assert IchitaroConverter.ICHITARO_EXTENSIONS == ('.jtd',)

    def test_is_cancelled_false(self, converter: IchitaroConverter):
        """キャンセルされていない状態のテスト"""
        assert not converter.is_cancelled()

    def test_wait_with_cancel_check_normal(self, converter: IchitaroConverter):
        """通常の待機テスト"""
        start_time = time.time()
        converter._wait_with_cancel_check(0.1)
        elapsed = time.time() - start_time
        assert elapsed >= 0.1

    def test_wait_with_cancel_check_cancelled(self):
        """キャンセル時の待機中断テスト"""
        cancel_flag = [False]

        def cancel_check():
            return cancel_flag[0]

        converter = IchitaroConverter(cancel_check=cancel_check)

        # 0.5秒後にキャンセル
        def delayed_cancel():
            time.sleep(0.1)
            cancel_flag[0] = True

        import threading
        cancel_thread = threading.Thread(target=delayed_cancel)
        cancel_thread.start()

        with patch.object(converter, '_cleanup_ichitaro_windows'):
            with pytest.raises(CancelledError):
                converter._wait_with_cancel_check(5.0)

        cancel_thread.join()

    def test_escape_for_send_keys_braces(self):
        """波括弧のエスケープテスト"""
        text = "test{value}text"
        escaped = IchitaroConverter._escape_for_send_keys(text)
        assert escaped == "test{{value}}text"

    def test_escape_for_send_keys_special_chars(self):
        """特殊文字のエスケープテスト"""
        text = "test+^%~()"
        escaped = IchitaroConverter._escape_for_send_keys(text)
        assert escaped == "test{+}{^}{%}{~}{(}{)}"

    def test_escape_for_send_keys_mixed(self):
        """混合特殊文字のエスケープテスト"""
        text = "C:\\Users\\test{file}.txt"
        escaped = IchitaroConverter._escape_for_send_keys(text)
        assert escaped == "C:\\Users\\test{{file}}.txt"

    def test_escape_for_send_keys_complex(self):
        """複雑なパスのエスケープテスト"""
        text = "C:\\Test (2026)\\file+doc.jtd"
        escaped = IchitaroConverter._escape_for_send_keys(text)
        assert escaped == "C:\\Test {(}2026{)}\\file{+}doc.jtd"

    @patch('converters.ichitaro_converter.os.startfile')
    @patch('converters.ichitaro_converter.Application')
    @patch('time.sleep')
    def test_open_ichitaro_file_success(
        self,
        mock_sleep: Mock,
        mock_app_class: Mock,
        mock_startfile: Mock,
        converter: IchitaroConverter,
        mock_jtd_file: Path
    ):
        """一太郎ファイルオープン成功テスト"""
        mock_app = MagicMock()
        mock_window = MagicMock()
        mock_window.window_text.return_value = "test.jtd - 一太郎"
        mock_app.top_window.return_value = mock_window
        mock_app_class.return_value.connect.return_value = mock_app

        app, window = converter._open_ichitaro_file(str(mock_jtd_file), max_wait=10)

        assert app is mock_app
        assert window is mock_window
        mock_startfile.assert_called_once_with(str(mock_jtd_file))
        mock_window.set_focus.assert_called_once()

    @patch('converters.ichitaro_converter.os.startfile')
    @patch('converters.ichitaro_converter.Application')
    @patch('time.sleep')
    def test_open_ichitaro_file_connection_failure(
        self,
        mock_sleep: Mock,
        mock_app_class: Mock,
        mock_startfile: Mock,
        converter: IchitaroConverter,
        mock_jtd_file: Path
    ):
        """一太郎接続失敗テスト"""
        mock_app_class.return_value.connect.side_effect = Exception("Connection failed")

        with patch.object(converter, '_cleanup_ichitaro_windows'):
            app, window = converter._open_ichitaro_file(str(mock_jtd_file), max_wait=10)

        assert app is None
        assert window is None

    @patch('converters.ichitaro_converter.os.startfile')
    @patch('time.sleep')
    def test_open_ichitaro_file_cancelled(
        self,
        mock_sleep: Mock,
        mock_startfile: Mock,
        mock_jtd_file: Path
    ):
        """一太郎オープン中のキャンセルテスト"""
        cancel_flag = [False]

        def cancel_check():
            return cancel_flag[0]

        converter = IchitaroConverter(cancel_check=cancel_check)
        cancel_flag[0] = True

        with patch.object(converter, '_cleanup_ichitaro_windows'):
            with pytest.raises(CancelledError):
                converter._open_ichitaro_file(str(mock_jtd_file), max_wait=10)

    @patch('converters.ichitaro_converter.send_keys')
    @patch('time.sleep')
    def test_execute_print_sequence_success(
        self,
        mock_sleep: Mock,
        mock_send_keys: Mock,
        converter: IchitaroConverter,
        temp_dir: Path
    ):
        """印刷シーケンス成功テスト"""
        output_path = temp_dir / "output.pdf"

        mock_app = MagicMock()
        mock_window = MagicMock()
        mock_dialog = MagicMock()
        mock_combo = MagicMock()
        mock_button = MagicMock()

        mock_app.top_window.return_value = mock_window
        mock_window.child_window.return_value = mock_dialog
        mock_dialog.child_window.side_effect = [mock_combo, mock_button]

        with patch.object(converter, '_handle_save_dialog'):
            result = converter._execute_print_sequence(mock_app, str(output_path))

        assert result is True
        mock_combo.select.assert_called_once_with("Microsoft Print to PDF")

    @patch('converters.ichitaro_converter.send_keys')
    @patch('time.sleep')
    def test_execute_print_sequence_printer_select_retry(
        self,
        mock_sleep: Mock,
        mock_send_keys: Mock,
        converter: IchitaroConverter,
        temp_dir: Path
    ):
        """プリンター選択リトライテスト"""
        output_path = temp_dir / "output.pdf"

        mock_app = MagicMock()
        mock_window = MagicMock()
        mock_dialog = MagicMock()
        mock_combo = MagicMock()

        mock_app.top_window.return_value = mock_window
        mock_window.child_window.return_value = mock_dialog

        # 最初は失敗、2回目で成功
        mock_dialog.child_window.side_effect = [
            Exception("Not ready"),
            mock_combo,
            MagicMock()  # ok_button
        ]

        with patch.object(converter, '_handle_save_dialog'):
            result = converter._execute_print_sequence(mock_app, str(output_path))

        assert result is True

    @patch('time.sleep')
    def test_close_ichitaro_success(
        self,
        mock_sleep: Mock,
        converter: IchitaroConverter
    ):
        """一太郎クローズ成功テスト"""
        mock_app = MagicMock()
        mock_window = MagicMock()
        mock_window.window_text.return_value = "test.jtd - 一太郎"

        converter._close_ichitaro(mock_app, mock_window)

        mock_app.kill.assert_called_once()

    def test_close_ichitaro_none_objects(self, converter: IchitaroConverter):
        """None オブジェクトのクローズテスト"""
        # 例外が発生しないことを確認
        converter._close_ichitaro(None, None)

    @patch('os.path.exists')
    @patch('os.path.getsize')
    @patch('time.sleep')
    def test_wait_for_output_file_success(
        self,
        mock_sleep: Mock,
        mock_getsize: Mock,
        mock_exists: Mock,
        converter: IchitaroConverter,
        temp_dir: Path
    ):
        """出力ファイル待機成功テスト"""
        output_path = temp_dir / "output.pdf"

        # ファイルが存在し、サイズが安定している
        mock_exists.return_value = True
        mock_getsize.return_value = 1024

        result = converter._wait_for_output_file(
            str(output_path),
            "test.jtd",
            save_wait=10
        )

        assert result == str(output_path)

    @patch('os.path.exists')
    @patch('time.sleep')
    def test_wait_for_output_file_timeout(
        self,
        mock_sleep: Mock,
        mock_exists: Mock,
        converter: IchitaroConverter,
        temp_dir: Path
    ):
        """出力ファイル待機タイムアウトテスト"""
        output_path = temp_dir / "output.pdf"
        mock_exists.return_value = False

        result = converter._wait_for_output_file(
            str(output_path),
            "test.jtd",
            save_wait=0.1
        )

        assert result is None

    @patch('converters.ichitaro_converter.Application')
    @patch('time.sleep')
    def test_cleanup_ichitaro_windows_success(
        self,
        mock_sleep: Mock,
        mock_app_class: Mock,
        converter: IchitaroConverter
    ):
        """一太郎ウィンドウクリーンアップ成功テスト"""
        mock_app = MagicMock()
        mock_app_class.return_value.connect.return_value = mock_app

        converter._cleanup_ichitaro_windows()

        mock_app.kill.assert_called_once()

    @patch('converters.ichitaro_converter.Application')
    def test_cleanup_ichitaro_windows_no_process(
        self,
        mock_app_class: Mock,
        converter: IchitaroConverter
    ):
        """一太郎プロセスなしのクリーンアップテスト"""
        mock_app_class.return_value.connect.side_effect = Exception("No process")

        # 例外が発生しないことを確認
        converter._cleanup_ichitaro_windows()

    @patch('converters.ichitaro_converter.os.path.exists')
    @patch('converters.ichitaro_converter.os.path.getsize')
    @patch('converters.ichitaro_converter.os.remove')
    @patch('converters.ichitaro_converter.os.startfile')
    @patch('converters.ichitaro_converter.Application')
    @patch('converters.ichitaro_converter.send_keys')
    @patch('time.sleep')
    def test_convert_success(
        self,
        mock_sleep: Mock,
        mock_send_keys: Mock,
        mock_app_class: Mock,
        mock_startfile: Mock,
        mock_remove: Mock,
        mock_getsize: Mock,
        mock_exists: Mock,
        converter: IchitaroConverter,
        mock_jtd_file: Path,
        temp_dir: Path
    ):
        """一太郎変換成功テスト（統合）"""
        output_path = temp_dir / "output.pdf"

        # モックの設定
        mock_app = MagicMock()
        mock_window = MagicMock()
        mock_window.window_text.return_value = "test.jtd - 一太郎"
        mock_app.top_window.return_value = mock_window
        mock_app_class.return_value.connect.return_value = mock_app

        # ファイル存在・サイズチェック
        mock_exists.return_value = True
        mock_getsize.return_value = 1024

        # ダイアログモック
        mock_dialog = MagicMock()
        mock_combo = MagicMock()
        mock_button = MagicMock()
        mock_window.child_window.return_value = mock_dialog
        mock_dialog.child_window.side_effect = [mock_combo, mock_button]

        with patch.object(converter, '_handle_save_dialog'):
            result = converter.convert(str(mock_jtd_file), str(output_path))

        assert result == str(output_path)

    @patch('converters.ichitaro_converter.os.startfile')
    @patch('converters.ichitaro_converter.Application')
    @patch('time.sleep')
    def test_convert_cancelled_during_open(
        self,
        mock_sleep: Mock,
        mock_app_class: Mock,
        mock_startfile: Mock,
        mock_jtd_file: Path,
        temp_dir: Path
    ):
        """一太郎変換キャンセルテスト"""
        output_path = temp_dir / "output.pdf"

        cancel_flag = [False]

        def cancel_check():
            return cancel_flag[0]

        converter = IchitaroConverter(cancel_check=cancel_check)

        # オープン時にキャンセル
        cancel_flag[0] = True

        with patch.object(converter, '_cleanup_ichitaro_windows'):
            with pytest.raises(CancelledError):
                converter.convert(str(mock_jtd_file), str(output_path))

    @patch('converters.ichitaro_converter.os.startfile')
    @patch('converters.ichitaro_converter.Application')
    @patch('time.sleep')
    def test_convert_system_exit_propagation(
        self,
        mock_sleep: Mock,
        mock_app_class: Mock,
        mock_startfile: Mock,
        converter: IchitaroConverter,
        mock_jtd_file: Path,
        temp_dir: Path
    ):
        """SystemExitの伝播テスト"""
        output_path = temp_dir / "output.pdf"
        mock_startfile.side_effect = SystemExit(1)

        with pytest.raises(SystemExit):
            converter.convert(str(mock_jtd_file), str(output_path))

    @patch('converters.ichitaro_converter.os.startfile')
    @patch('converters.ichitaro_converter.Application')
    @patch('time.sleep')
    def test_convert_keyboard_interrupt_propagation(
        self,
        mock_sleep: Mock,
        mock_app_class: Mock,
        mock_startfile: Mock,
        converter: IchitaroConverter,
        mock_jtd_file: Path,
        temp_dir: Path
    ):
        """KeyboardInterruptの伝播テスト"""
        output_path = temp_dir / "output.pdf"
        mock_startfile.side_effect = KeyboardInterrupt()

        with pytest.raises(KeyboardInterrupt):
            converter.convert(str(mock_jtd_file), str(output_path))


if __name__ == "__main__":
    pytest.main([__file__, "-v", "--tb=short"])
