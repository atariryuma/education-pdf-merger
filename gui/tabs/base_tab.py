"""
ベースタブクラス

全てのタブで共有される基本機能を提供
"""
import logging
import tkinter as tk
from tkinter import ttk, scrolledtext
from pathlib import Path
from tkinter import messagebox
from typing import Any, List, Optional, Callable, Tuple, TYPE_CHECKING

from gui.styles import COLORS, FONTS
from gui.utils import log_message
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

    def __init__(self, notebook: ttk.Notebook, config: "ConfigLoader", status_bar: tk.Label) -> None:
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

    def create_log_frame(self, height: int = 10, parent=None) -> None:
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
        tk.Label(log_frame, text="実行ログ:", font=FONTS['default_bold']).pack(anchor="w", pady=(0, 5))
        self.log_widget = scrolledtext.ScrolledText(
            log_frame, width=80, height=height, state="disabled", wrap=tk.WORD
        )
        self.log_widget.pack(fill="both", expand=True)

    def setup_gui_logging(self, logger_names: list = None) -> None:
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
            lambda msg, msg_type: log_message(self.log_widget, msg, msg_type) if self.log_widget else None
        )
        self._gui_handler.setLevel(logging.INFO)
        formatter = logging.Formatter('%(message)s')
        self._gui_handler.setFormatter(formatter)

        # 各ロガーにハンドラを追加
        for name in logger_names:
            logger = logging.getLogger(name)
            # 重複防止: 既にGUILogHandlerがあれば追加しない
            has_gui_handler = any(
                isinstance(h, GUILogHandler) for h in logger.handlers
            )
            if not has_gui_handler:
                logger.addHandler(self._gui_handler)

    def remove_gui_logging(self) -> None:
        """GUIログハンドラを削除"""
        if hasattr(self, '_gui_handler') and self._gui_handler:
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
        error_title: str = "パスエラー"
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

    def create_collapsible_section(
        self,
        parent: tk.Widget,
        title_collapsed: str,
        title_expanded: str,
        **pack_kwargs: Any
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
            font=FONTS['default_bold'],
            relief="flat",
            anchor="w",
            cursor="hand2",
            bg=COLORS['surface_dim']
        )
        toggle_btn.pack(**pack_kwargs) if pack_kwargs else toggle_btn.pack(fill="x")

        return toggle_btn, content_frame
