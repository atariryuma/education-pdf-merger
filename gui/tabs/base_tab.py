"""
ベースタブクラス

全てのタブで共有される基本機能を提供
"""
import logging
import threading
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog
from pathlib import Path
from tkinter import messagebox
from typing import Any, List, Optional, Callable, Tuple, TYPE_CHECKING

from gui.styles import COLORS, FONTS
from gui.utils import log_message, set_button_state, thread_safe_call
from infrastructure.com_utils import com_apartment
from infrastructure.path_validator import PathValidator

if TYPE_CHECKING:
    from infrastructure.config_loader import ConfigLoader


class GUILogHandler(logging.Handler):
    """GUIのログウィジェットに出力するログハンドラ"""

    def __init__(self, log_callback: Callable[[str, str], None]) -> None:
        """
        Args:
            log_callback: ログメッセージを受け取るコールバック関数
                         引数: (message: str, msg_type: str)
        """
        super().__init__()
        self.log_callback = log_callback

    def emit(self, record: logging.LogRecord) -> None:
        """ログレコードを処理"""
        try:
            msg = self.format(record)
            # ログレベルに応じてメッセージタイプを決定
            if record.levelno >= logging.ERROR:
                msg_type = "error"
            elif record.levelno >= logging.WARNING:
                msg_type = "warning"
            elif record.levelno <= logging.DEBUG:
                msg_type = "normal"
            else:
                msg_type = "info"
            self.log_callback(msg, msg_type)
        except Exception:
            self.handleError(record)


class BaseTab:
    """タブの基底クラス"""

    def __init__(
        self, notebook: ttk.Notebook, config: "ConfigLoader", status_bar: tk.Label
    ) -> None:
        """
        Args:
            notebook: タブを追加するNotebookウィジェット
            config: ConfigLoaderインスタンス
            status_bar: ステータスバーのLabelウィジェット
        """
        self.notebook = notebook
        self.config = config
        self.status_bar = status_bar
        self.tab = ttk.Frame(notebook)
        self.log_widget: Optional[scrolledtext.ScrolledText] = None

    def add_to_notebook(self, text: str) -> None:
        """タブをNotebookに追加"""
        self.notebook.add(self.tab, text=text)

    def create_log_frame(
        self, height: int = 10, parent: Optional[tk.Widget] = None
    ) -> None:
        """
        ログフレームを作成

        Args:
            height: ログウィジェットの高さ（行数）
            parent: 親ウィジェット（省略時はタブ自体）
        """
        if parent is None:
            parent = self.tab
        log_frame = tk.Frame(parent)
        log_frame.pack(fill="both", expand=True, padx=20, pady=(5, 15))
        tk.Label(log_frame, text="実行ログ:", font=FONTS["default_bold"]).pack(
            anchor="w", pady=(0, 5)
        )
        self.log_widget = scrolledtext.ScrolledText(
            log_frame, width=80, height=height, state="disabled", wrap=tk.WORD
        )
        self.log_widget.pack(fill="both", expand=True)

    def setup_gui_logging(self, logger_names: Optional[List[str]] = None) -> None:
        """
        ロガーにGUIハンドラを追加して、ログをGUIに表示する

        Args:
            logger_names: ハンドラを追加するロガー名のリスト
                         省略時は主要モジュールのロガーに追加
        """
        if self.log_widget is None:
            return

        if logger_names is None:
            from shared.constants import AppConstants

            logger_names = AppConstants.GUI_LOGGER_NAMES

        # GUIログハンドラを作成
        self._gui_handler = GUILogHandler(
            lambda msg, msg_type: log_message(self.log_widget, msg, msg_type)
            if self.log_widget
            else None
        )
        self._gui_handler.setLevel(logging.INFO)
        formatter = logging.Formatter("%(message)s")
        self._gui_handler.setFormatter(formatter)

        # 各ロガーにハンドラを追加
        for name in logger_names:
            logger = logging.getLogger(name)
            # 重複防止: 既にGUILogHandlerがあれば追加しない
            has_gui_handler = any(isinstance(h, GUILogHandler) for h in logger.handlers)
            if not has_gui_handler:
                logger.addHandler(self._gui_handler)

    def remove_gui_logging(self) -> None:
        """GUIログハンドラを削除"""
        if hasattr(self, "_gui_handler") and self._gui_handler:
            from shared.constants import AppConstants

            for name in AppConstants.GUI_LOGGER_NAMES:
                logger = logging.getLogger(name)
                if self._gui_handler in logger.handlers:
                    logger.removeHandler(self._gui_handler)
            self._gui_handler = None

    def log(self, message: str, msg_type: str = "info") -> None:
        """
        ログウィジェットにメッセージを出力

        Args:
            message: ログメッセージ
            msg_type: メッセージタイプ ("info", "success", "warning", "error", "normal")
        """
        if self.log_widget:
            log_message(self.log_widget, message, msg_type)

    def update_status(self, message: str) -> None:
        """
        ステータスメッセージを更新（ログに出力）

        Args:
            message: ステータスメッセージ
        """
        self.log(message, "info")

    def validate_path(
        self,
        path: str,
        path_type: str = "directory",
        must_exist: bool = True,
        allowed_extensions: Optional[List[str]] = None,
        error_title: str = "パスエラー",
    ) -> Optional[Path]:
        """
        パスを検証し、無効な場合はエラーダイアログを表示

        Args:
            path: 検証するパス文字列
            path_type: "directory" または "file"
            must_exist: 存在を要求するか
            allowed_extensions: 許可する拡張子リスト（fileのみ）
            error_title: エラーダイアログのタイトル

        Returns:
            Optional[Path]: 検証済みパス。無効な場合はNone（エラーダイアログ表示済み）
        """
        if path_type == "directory":
            is_valid, error_msg, validated_path = PathValidator.validate_directory(
                path, must_exist=must_exist
            )
        else:
            is_valid, error_msg, validated_path = PathValidator.validate_file_path(
                path, must_exist=must_exist, allowed_extensions=allowed_extensions
            )

        if is_valid and validated_path:
            return validated_path

        messagebox.showerror(error_title, error_msg or "パスが無効です")
        return None

    def ask_folder(
        self,
        title: str = "フォルダを選択",
        initial_dir: Optional[str] = None,
    ) -> Optional[Path]:
        """
        フォルダ選択ダイアログを表示し、検証済みPathを返す

        - 検証エラー時はメッセージダイアログを表示
        - ダイアログのキャンセル時はNoneを返す
        - 例外時もメッセージダイアログを表示してNoneを返す

        Args:
            title: ダイアログのタイトル
            initial_dir: 初期表示ディレクトリ

        Returns:
            検証済みPath、またはNone
        """
        try:
            kwargs: dict = {"title": title}
            if initial_dir:
                kwargs["initialdir"] = initial_dir
            directory = filedialog.askdirectory(**kwargs)
            if not directory:
                return None
            return self.validate_path(directory, "directory")
        except Exception as e:
            messagebox.showerror(
                "参照エラー",
                f"フォルダの参照中にエラーが発生しました。\n\n詳細: {e}",
            )
            return None

    def ask_file_open(
        self,
        title: str = "ファイルを選択",
        filetypes: Optional[List[Tuple[str, str]]] = None,
        initial_dir: Optional[str] = None,
        allowed_extensions: Optional[List[str]] = None,
    ) -> Optional[Path]:
        """
        既存ファイル選択ダイアログを表示し、検証済みPathを返す

        Args:
            title: ダイアログのタイトル
            filetypes: filedialog 用のファイルタイプ一覧
            initial_dir: 初期表示ディレクトリ
            allowed_extensions: 許可する拡張子（検証用）

        Returns:
            検証済みPath、またはNone
        """
        try:
            kwargs: dict = {"title": title}
            if filetypes:
                kwargs["filetypes"] = filetypes
            if initial_dir:
                kwargs["initialdir"] = initial_dir
            file_path = filedialog.askopenfilename(**kwargs)
            if not file_path:
                return None
            return self.validate_path(
                file_path,
                "file",
                must_exist=True,
                allowed_extensions=allowed_extensions,
            )
        except Exception as e:
            messagebox.showerror(
                "参照エラー",
                f"ファイルの参照中にエラーが発生しました。\n\n詳細: {e}",
            )
            return None

    def ask_file_save(
        self,
        title: str = "保存先を選択",
        filetypes: Optional[List[Tuple[str, str]]] = None,
        initial_dir: Optional[str] = None,
        initial_file: str = "",
        default_extension: str = "",
        allowed_extensions: Optional[List[str]] = None,
    ) -> Optional[Path]:
        """
        ファイル保存ダイアログを表示し、検証済みPathを返す

        Args:
            title: ダイアログのタイトル
            filetypes: filedialog 用のファイルタイプ一覧
            initial_dir: 初期表示ディレクトリ
            initial_file: 初期ファイル名
            default_extension: デフォルト拡張子
            allowed_extensions: 許可する拡張子（検証用）

        Returns:
            検証済みPath、またはNone
        """
        try:
            kwargs: dict = {"title": title}
            if filetypes:
                kwargs["filetypes"] = filetypes
            if initial_dir:
                kwargs["initialdir"] = initial_dir
            if initial_file:
                kwargs["initialfile"] = initial_file
            if default_extension:
                kwargs["defaultextension"] = default_extension
            file_path = filedialog.asksaveasfilename(**kwargs)
            if not file_path:
                return None
            return self.validate_path(
                file_path,
                "file",
                must_exist=False,
                allowed_extensions=allowed_extensions,
            )
        except Exception as e:
            messagebox.showerror(
                "参照エラー",
                f"ファイルの参照中にエラーが発生しました。\n\n詳細: {e}",
            )
            return None

    def poll_thread(
        self,
        thread: threading.Thread,
        on_complete: Callable[[], None],
        timeout_seconds: float = 10.0,
        poll_interval_ms: int = 200,
        on_timeout: Optional[Callable[[], None]] = None,
    ) -> None:
        """
        スレッドの完了を非ブロッキングでポーリングし、完了時にコールバックを呼ぶ

        UIフリーズを避けるため、`tab.after()` で再帰的に自身をスケジュールする。

        Args:
            thread: 監視するスレッド
            on_complete: スレッド完了時に呼ばれるコールバック
            timeout_seconds: タイムアウト秒数
            poll_interval_ms: ポーリング間隔（ミリ秒）
            on_timeout: タイムアウト時のコールバック（省略時は何もしない）
        """
        max_polls = max(1, int((timeout_seconds * 1000) / poll_interval_ms))
        state = {"count": 0}

        def _tick() -> None:
            if not thread.is_alive():
                on_complete()
                return
            state["count"] += 1
            if state["count"] >= max_polls:
                if on_timeout is not None:
                    on_timeout()
                return
            self.tab.after(poll_interval_ms, _tick)

        self.tab.after(poll_interval_ms, _tick)

    def run_in_thread(
        self,
        target: Callable[[], None],
        button: Optional[tk.Widget] = None,
        running_status: str = "🔄 実行中...",
        error_title: str = "実行エラー",
        on_error: Optional[Callable[[Exception], None]] = None,
        on_finally: Optional[Callable[[], None]] = None,
        com_sta: Optional[bool] = None,
    ) -> threading.Thread:
        """
        バックグラウンドスレッドで処理を実行する共通ヘルパー

        - 実行前にボタンを無効化、終了後に再有効化
        - 例外時はログ・ステータス・エラーダイアログを統一表示
        - `com_sta` が指定されていれば `com_apartment` で囲んで実行
        - `on_finally` で実行終了時の追加処理（フラグリセット等）を指定可能

        Args:
            target: 実行する関数（引数なし）
            button: 実行中に無効化するボタン
            running_status: 実行中ボタンに表示するテキスト
            error_title: エラーダイアログのタイトル
            on_error: 追加のエラーハンドラ（オプション）
            on_finally: スレッド終了時に必ず呼ばれるコールバック（オプション）
            com_sta: COMアパートメント (True=STA, False=MTA, None=囲まない)

        Returns:
            起動したスレッド（daemon）
        """
        if button is not None:
            set_button_state(
                button, False, getattr(self, "status_label", None), running_status
            )

        def _run_target() -> None:
            if com_sta is None:
                target()
            else:
                with com_apartment(sta=com_sta):
                    target()

        def _wrapper() -> None:
            try:
                _run_target()
            except Exception as e:
                error_msg = str(e)
                self.log(f"❌ エラー: {error_msg}", "error")
                self.update_status("❌ エラーが発生しました")
                thread_safe_call(
                    self.tab,
                    lambda: messagebox.showerror(
                        error_title, f"エラーが発生しました。\n\n詳細:\n{error_msg}"
                    ),
                )
                if on_error is not None:
                    on_error(e)
            finally:
                if button is not None:
                    set_button_state(
                        button, True, getattr(self, "status_label", None), ""
                    )
                if on_finally is not None:
                    on_finally()

        thread = threading.Thread(target=_wrapper, daemon=True)
        thread.start()
        return thread

    def create_collapsible_section(
        self,
        parent: tk.Widget,
        title_collapsed: str,
        title_expanded: str,
        **pack_kwargs: Any,
    ) -> Tuple[tk.Button, tk.Frame]:
        """
        折りたたみ可能なセクションを作成

        Args:
            parent: 親ウィジェット
            title_collapsed: 折りたたみ時のボタンテキスト（例: "▶ 使い方を表示"）
            title_expanded: 展開時のボタンテキスト（例: "▼ 使い方を非表示"）
            **pack_kwargs: toggle_button.pack() に渡す引数

        Returns:
            Tuple[tk.Button, tk.Frame]: (トグルボタン, コンテンツフレーム)
            コンテンツフレームに子ウィジェットを追加してください。
            デフォルトでは折りたたまれています。
        """
        content_frame = tk.Frame(parent)

        def toggle() -> None:
            if content_frame.winfo_manager():  # currently packed
                content_frame.pack_forget()
                toggle_btn.config(text=title_collapsed)
            else:
                content_frame.pack(fill="x", after=toggle_btn)
                toggle_btn.config(text=title_expanded)

        toggle_btn = tk.Button(
            parent,
            text=title_collapsed,
            command=toggle,
            font=FONTS["default_bold"],
            relief="flat",
            anchor="w",
            cursor="hand2",
            bg=COLORS["surface_dim"],
        )
        toggle_btn.pack(**pack_kwargs) if pack_kwargs else toggle_btn.pack(fill="x")

        return toggle_btn, content_frame
