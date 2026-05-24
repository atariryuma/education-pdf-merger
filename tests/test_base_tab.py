"""BaseTab のヘルパーメソッドのユニットテスト"""
import threading
import time
import tkinter as tk
from tkinter import ttk
from unittest.mock import MagicMock, patch

import pytest


@pytest.fixture
def tk_root():
    """tkinterのルートウィンドウ"""
    try:
        root = tk.Tk()
        root.withdraw()
    except tk.TclError as e:
        pytest.skip(f"tkinter利用不可: {e}")
    yield root
    try:
        root.destroy()
    except tk.TclError:
        pass


@pytest.fixture
def base_tab(tk_root):
    """テスト用BaseTabインスタンス"""
    from gui.tabs.base_tab import BaseTab
    notebook = ttk.Notebook(tk_root)
    config = MagicMock()
    status_bar = tk.Label(tk_root)
    return BaseTab(notebook, config, status_bar)


@pytest.mark.unit
class TestAskFolder:
    """ask_folder のテスト"""

    @patch("gui.tabs.base_tab.filedialog")
    def test_returns_validated_path_when_selected(self, mock_filedialog, base_tab, tmp_path):
        mock_filedialog.askdirectory.return_value = str(tmp_path)
        result = base_tab.ask_folder(title="テスト")
        assert result is not None
        assert str(result) == str(tmp_path)

    @patch("gui.tabs.base_tab.filedialog")
    def test_returns_none_on_cancel(self, mock_filedialog, base_tab):
        mock_filedialog.askdirectory.return_value = ""
        result = base_tab.ask_folder()
        assert result is None

    @patch("gui.tabs.base_tab.filedialog")
    @patch("gui.tabs.base_tab.messagebox")
    def test_returns_none_on_exception(self, mock_msg, mock_filedialog, base_tab):
        mock_filedialog.askdirectory.side_effect = RuntimeError("test")
        result = base_tab.ask_folder()
        assert result is None
        mock_msg.showerror.assert_called_once()


@pytest.mark.unit
class TestAskFileOpen:
    """ask_file_open のテスト"""

    @patch("gui.tabs.base_tab.filedialog")
    def test_returns_validated_path(self, mock_filedialog, base_tab, tmp_path):
        f = tmp_path / "test.txt"
        f.write_text("x")
        mock_filedialog.askopenfilename.return_value = str(f)
        result = base_tab.ask_file_open()
        assert result is not None

    @patch("gui.tabs.base_tab.filedialog")
    def test_cancel_returns_none(self, mock_filedialog, base_tab):
        mock_filedialog.askopenfilename.return_value = ""
        assert base_tab.ask_file_open() is None


@pytest.mark.unit
class TestAskFileSave:
    """ask_file_save のテスト"""

    @patch("gui.tabs.base_tab.filedialog")
    def test_returns_validated_path_for_new_file(self, mock_filedialog, base_tab, tmp_path):
        new_file = tmp_path / "new.pdf"
        mock_filedialog.asksaveasfilename.return_value = str(new_file)
        result = base_tab.ask_file_save(allowed_extensions=[".pdf"])
        assert result is not None
        assert str(result).endswith(".pdf")

    @patch("gui.tabs.base_tab.filedialog")
    def test_cancel_returns_none(self, mock_filedialog, base_tab):
        mock_filedialog.asksaveasfilename.return_value = ""
        assert base_tab.ask_file_save() is None


@pytest.mark.unit
class TestPollThread:
    """poll_thread のテスト"""

    def test_calls_on_complete_when_thread_done(self, base_tab, tk_root):
        completed = threading.Event()

        def quick_task():
            time.sleep(0.05)

        thread = threading.Thread(target=quick_task, daemon=True)
        thread.start()

        def on_complete():
            completed.set()

        base_tab.poll_thread(thread, on_complete=on_complete, timeout_seconds=2.0, poll_interval_ms=50)
        # tkinterイベントループを少し回す
        for _ in range(50):
            tk_root.update()
            if completed.is_set():
                break
            time.sleep(0.05)
        assert completed.is_set()

    def test_calls_on_timeout_when_thread_too_slow(self, base_tab, tk_root):
        timed_out = threading.Event()
        stop_signal = threading.Event()

        def slow_task():
            stop_signal.wait(timeout=5.0)

        thread = threading.Thread(target=slow_task, daemon=True)
        thread.start()

        def on_complete():
            pass

        def on_timeout():
            timed_out.set()

        base_tab.poll_thread(
            thread,
            on_complete=on_complete,
            timeout_seconds=0.3,
            poll_interval_ms=50,
            on_timeout=on_timeout,
        )
        # タイムアウトを待つ
        for _ in range(50):
            tk_root.update()
            if timed_out.is_set():
                break
            time.sleep(0.05)
        stop_signal.set()  # スレッド終了
        assert timed_out.is_set()


@pytest.mark.unit
class TestRunInThread:
    """run_in_thread のテスト"""

    def test_runs_target_in_background(self, base_tab):
        called = threading.Event()

        def target():
            called.set()

        thread = base_tab.run_in_thread(target)
        thread.join(timeout=2.0)
        assert called.is_set()

    @patch("gui.tabs.base_tab.thread_safe_call")
    @patch("gui.tabs.base_tab.messagebox")
    def test_shows_error_dialog_on_exception(
        self, mock_msg, mock_tsc, base_tab
    ):
        """例外発生時にエラーダイアログを表示する（thread_safe_callを介して即実行）"""
        # thread_safe_call が渡された関数を即座に呼ぶように動作させる
        mock_tsc.side_effect = lambda widget, func: func()

        def failing_target():
            raise ValueError("テストエラー")

        thread = base_tab.run_in_thread(failing_target, error_title="テスト")
        thread.join(timeout=2.0)

        mock_msg.showerror.assert_called_once()
        args, _kwargs = mock_msg.showerror.call_args
        assert args[0] == "テスト"
        assert "テストエラー" in args[1]
