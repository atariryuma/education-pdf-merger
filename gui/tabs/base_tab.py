"""
ベースタブクラス

全てのタブで共有される基本機能を提供
"""
import logging
import tkinter as tk
from tkinter import ttk, scrolledtext
from typing import Optional, Callable, Tuple, TYPE_CHECKING

from gui.utils import log_message

if TYPE_CHECKING:
    from config_loader import ConfigLoader


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
        tk.Label(log_frame, text="実行ログ:", font=("メイリオ", 10, "bold")).pack(anchor="w", pady=(0, 5))
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
            from constants import AppConstants
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
            from constants import AppConstants

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

    def create_scrollable_container(self) -> Tuple[tk.Canvas, tk.Scrollbar, tk.Frame]:
        """
        スクロール可能なコンテナを作成（全タブ共通処理）

        このメソッドは、pdf_tab、settings_tab、excel_tabで重複していた
        スクロール処理を統一するために作成されました。

        Returns:
            Tuple[tk.Canvas, tk.Scrollbar, tk.Frame]: (canvas, scrollbar, scrollable_frame)

        Note:
            マウスホイール処理は自動的にバインドされます。
            返却されたscrollable_frameに子ウィジェットを追加してください。
        """
        # スクロール可能なメインコンテナ
        canvas = tk.Canvas(self.tab, highlightthickness=0, bg="#f0f0f0")
        scrollbar = tk.Scrollbar(self.tab, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="#f0f0f0")

        # scrollregionを更新する関数
        def update_scrollregion(event: Optional[tk.Event] = None) -> None:
            canvas.update_idletasks()
            canvas.configure(scrollregion=canvas.bbox("all"))

        scrollable_frame.bind("<Configure>", update_scrollregion)

        # create_windowでウィンドウIDを保存
        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Canvasのサイズに合わせてscrollable_frameの幅を調整
        def on_canvas_configure(event: tk.Event) -> None:
            canvas.itemconfig(canvas_window, width=event.width)

        # マウスホイールでのスクロールを有効化
        def on_mousewheel(event: tk.Event) -> str:
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            return "break"

        # Canvas自体とすべての子ウィジェットにマウスホイールイベントをバインド
        def bind_mousewheel_recursive(widget: tk.Widget) -> None:
            widget.bind("<MouseWheel>", on_mousewheel)
            for child in widget.winfo_children():
                bind_mousewheel_recursive(child)

        canvas.bind("<Configure>", on_canvas_configure)
        canvas.bind("<MouseWheel>", on_mousewheel)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        # マウスホイール処理を遅延初期化（パフォーマンス向上）
        def deferred_mousewheel_bind() -> None:
            bind_mousewheel_recursive(scrollable_frame)

        # 100ms後にバインド（UI構築完了後）
        self.tab.after(100, deferred_mousewheel_bind)

        return canvas, scrollbar, scrollable_frame
