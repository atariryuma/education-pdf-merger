"""
GUIユーティリティモジュール

共通のUI操作やヘルパー関数を提供
"""
import logging
import os
import subprocess
import sys
import tkinter as tk
from datetime import datetime
from typing import Any, Optional, Callable

from gui.styles import COLORS, FONTS, BUTTON_STYLES

logger = logging.getLogger(__name__)


class ToolTip:
    """シンプルなツールチップ（マウスホバー時の説明）"""

    def __init__(self, widget: tk.Widget, text: str, delay: int = 500) -> None:
        self.widget = widget
        self.text = text
        self.delay = delay
        self.tip_window: Optional[tk.Toplevel] = None
        self.id: Optional[str] = None

        self.widget.bind("<Enter>", self._on_enter)
        self.widget.bind("<Leave>", self._on_leave)
        self.widget.bind("<Button>", self._on_leave)

    def _on_enter(self, event: Optional[tk.Event] = None) -> None:
        self._cancel()
        self.id = self.widget.after(self.delay, self._show)

    def _on_leave(self, event: Optional[tk.Event] = None) -> None:
        self._cancel()
        self._hide()

    def _cancel(self) -> None:
        if self.id:
            self.widget.after_cancel(self.id)
            self.id = None

    def _show(self) -> None:
        if self.tip_window or not self.text:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tk.Label(
            tw,
            text=self.text,
            justify=tk.LEFT,
            background=COLORS["tooltip_bg"],
            foreground=COLORS["tooltip_fg"],
            relief=tk.SOLID,
            borderwidth=1,
            font=FONTS["small"],
            padx=8,
            pady=4,
        ).pack()

    def _hide(self) -> None:
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None


def center_window(window: tk.Toplevel) -> None:
    """ウィンドウを画面中央に配置"""
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")


def create_tooltip(widget: tk.Widget, text: str) -> ToolTip:
    """ウィジェットにツールチップを追加"""
    if hasattr(widget, "_tooltip") and widget._tooltip:
        widget._tooltip._cancel()
        widget._tooltip._hide()
    tooltip = ToolTip(widget, text)
    widget._tooltip = tooltip
    return tooltip


def thread_safe_call(widget: tk.Widget, func: Callable[[], Any]) -> None:
    """スレッドセーフにGUI更新を行う"""
    try:
        widget.after_idle(func)
    except tk.TclError:
        pass


def update_status(status_bar: tk.Label, message: str) -> None:
    """ステータスバーを更新（スレッドセーフ）"""

    def _update() -> None:
        try:
            timestamp = datetime.now().strftime("%H:%M:%S")
            status_bar.config(text=f"[{timestamp}] {message}")
        except tk.TclError:
            pass

    thread_safe_call(status_bar, _update)


def log_message(log_widget: Any, message: str, msg_type: str = "normal") -> None:
    """ログにメッセージを追加（スレッドセーフ）"""

    def _log() -> None:
        try:
            log_widget.config(state="normal")
            timestamp = datetime.now().strftime("%H:%M:%S")
            log_widget.insert(tk.END, f"[{timestamp}] {message}\n")
            log_widget.see(tk.END)
            log_widget.config(state="disabled")
        except tk.TclError:
            pass

    thread_safe_call(log_widget, _log)


def set_button_state(
    button: tk.Button,
    enabled: bool,
    status_label: Optional[tk.Label] = None,
    status_text: str = "",
) -> None:
    """ボタンの状態を変更（スレッドセーフ）"""

    def _set_state() -> None:
        try:
            if enabled:
                button.config(state="normal", cursor="hand2")
                if status_label:
                    status_label.config(text="")
            else:
                button.config(state="disabled", cursor="")
                if status_label:
                    status_label.config(text=status_text, fg="orange")
        except tk.TclError:
            pass

    thread_safe_call(button, _set_state)


def create_hover_button(
    parent: tk.Widget, text: str, command: Any, color: str = "primary", **kwargs: Any
) -> tk.Button:
    """ホバー効果付きボタンを作成"""
    style = BUTTON_STYLES.get(color, BUTTON_STYLES["primary"])
    bg_color = style["bg"]
    hover_color = style["activebackground"]

    button = tk.Button(
        parent,
        text=text,
        command=command,
        bg=bg_color,
        fg=style["fg"],
        cursor="hand2",
        relief=tk.RAISED,
        borderwidth=2,
        **kwargs,
    )
    button.bind("<Enter>", lambda e: button.config(bg=hover_color))
    button.bind("<Leave>", lambda e: button.config(bg=bg_color))
    return button


def open_file_or_folder(
    path: str, on_error: Optional[Callable[[str], None]] = None
) -> bool:
    """
    ファイルまたはフォルダをデフォルトアプリケーションで開く（Windows専用）

    Args:
        path: 開くファイルまたはフォルダのパス
        on_error: エラー時のコールバック関数
    """
    if sys.platform != "win32":
        if on_error:
            on_error("この機能はWindows専用です。")
        return False

    try:
        if not os.path.exists(path):
            if on_error:
                on_error(f"指定されたパスが存在しません:\n{path}")
            return False

        if os.path.isdir(path):
            try:
                subprocess.run(
                    ["explorer", path],
                    shell=False,
                    stdin=subprocess.DEVNULL,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    check=False,
                    timeout=5,
                )
            except subprocess.TimeoutExpired:
                pass  # エクスプローラーは非同期起動のためタイムアウトは正常
        else:
            os.startfile(path)
        return True

    except Exception as e:
        logger.error(f"ファイル/フォルダを開けません: {path}, エラー: {e}")
        if on_error:
            on_error(f"ファイル/フォルダを開けませんでした:\n{path}")
        return False
