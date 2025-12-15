"""
GUIユーティリティモジュール

共通のUI操作やヘルパー関数を提供
"""
import tkinter as tk
from datetime import datetime
from typing import Any, Optional, Callable


def thread_safe_call(widget: tk.Widget, func: Callable[[], Any]) -> None:
    """
    スレッドセーフにGUI更新を行う

    Args:
        widget: 任意のtkinterウィジェット（rootにアクセスするため）
        func: 実行する関数
    """
    try:
        # after_idleを使用してメインスレッドで実行
        widget.after_idle(func)
    except tk.TclError:
        # ウィジェットが破棄されている場合は無視
        pass


def update_status(status_bar: tk.Label, message: str) -> None:
    """
    ステータスバーを更新（スレッドセーフ）

    Args:
        status_bar: ステータスバーのLabelウィジェット
        message: 表示するメッセージ
    """
    def _update():
        try:
            timestamp = datetime.now().strftime("%H:%M:%S")
            status_bar.config(text=f"[{timestamp}] {message}")
        except tk.TclError:
            pass  # ウィジェットが破棄されている場合は無視

    thread_safe_call(status_bar, _update)


def log_message(log_widget: Any, message: str, msg_type: str = "normal") -> None:
    """
    ログにメッセージを追加（色付き、スレッドセーフ）

    Args:
        log_widget: ログ表示用のScrolledTextウィジェット
        message: 表示するメッセージ
        msg_type: メッセージタイプ（info, success, error, warning, normal）
    """
    def _log():
        try:
            log_widget.config(state="normal")

            # メッセージタイプに応じた装飾
            prefixes = {
                "info": "ℹ️ ",
                "success": "✅ ",
                "error": "❌ ",
                "warning": "⚠️ ",
            }
            prefix = prefixes.get(msg_type, "")

            timestamp = datetime.now().strftime("%H:%M:%S")
            log_widget.insert(tk.END, f"[{timestamp}] {prefix}{message}\n")
            log_widget.see(tk.END)
            log_widget.config(state="disabled")
        except tk.TclError:
            pass  # ウィジェットが破棄されている場合は無視

    thread_safe_call(log_widget, _log)


def set_button_state(
    button: tk.Button,
    enabled: bool,
    status_label: Optional[tk.Label] = None,
    status_text: str = ""
) -> None:
    """
    ボタンの状態を変更（スレッドセーフ）

    Args:
        button: 対象のボタン
        enabled: 有効にするかどうか
        status_label: ステータスラベル（オプション）
        status_text: ステータステキスト
    """
    def _set_state():
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
            pass  # ウィジェットが破棄されている場合は無視

    thread_safe_call(button, _set_state)


def create_hover_button(
    parent: tk.Widget,
    text: str,
    command: Any,
    color: str = "primary",
    **kwargs
) -> tk.Button:
    """
    ホバー効果付きボタンを作成

    Args:
        parent: 親ウィジェット
        text: ボタンテキスト
        command: コマンド
        color: カラータイプ（primary, secondary, warning, error）
        **kwargs: 追加のボタン引数

    Returns:
        tk.Button: 作成されたボタン
    """
    from gui.styles import BUTTON_STYLES

    style = BUTTON_STYLES.get(color, BUTTON_STYLES['primary'])
    bg_color = style['bg']
    hover_color = style['activebackground']

    button = tk.Button(
        parent,
        text=text,
        command=command,
        bg=bg_color,
        fg=style['fg'],
        cursor="hand2",
        relief=tk.RAISED,
        borderwidth=2,
        **kwargs
    )

    button.bind("<Enter>", lambda e: button.config(bg=hover_color))
    button.bind("<Leave>", lambda e: button.config(bg=bg_color))

    return button
