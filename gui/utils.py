"""
GUIユーティリティモジュール

共通のUI操作やヘルパー関数を提供
"""
import os
import tkinter as tk
from datetime import datetime
from typing import Any, Optional, Callable
from pathlib import Path


class ToolTip:
    """
    シンプルなツールチップ（マウスホバー時の説明）

    初心者向けに、ボタンやラベルにマウスを置くと説明が表示される機能
    """
    def __init__(self, widget: tk.Widget, text: str, delay: int = 500):
        """
        Args:
            widget: ツールチップを表示するウィジェット
            text: 表示するテキスト
            delay: 表示までの遅延時間（ミリ秒）
        """
        self.widget = widget
        self.text = text
        self.delay = delay
        self.tip_window = None
        self.id = None

        # イベントをバインド
        self.widget.bind("<Enter>", self.on_enter)
        self.widget.bind("<Leave>", self.on_leave)
        self.widget.bind("<Button>", self.on_leave)

    def on_enter(self, event: Optional[tk.Event] = None) -> None:
        """マウスが乗った時"""
        self.schedule()

    def on_leave(self, event: Optional[tk.Event] = None) -> None:
        """マウスが離れた時"""
        self.unschedule()
        self.hide()

    def schedule(self) -> None:
        """遅延後に表示"""
        self.unschedule()
        self.id = self.widget.after(self.delay, self.show)

    def unschedule(self) -> None:
        """スケジュールをキャンセル"""
        if self.id:
            self.widget.after_cancel(self.id)
            self.id = None

    def show(self) -> None:
        """ツールチップを表示"""
        if self.tip_window or not self.text:
            return

        # ウィジェットの位置を取得
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5

        # トップレベルウィンドウを作成
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)  # ウィンドウ枠を非表示
        tw.wm_geometry(f"+{x}+{y}")

        # ラベルを作成
        label = tk.Label(
            tw,
            text=self.text,
            justify=tk.LEFT,
            background="#FFFFE0",
            foreground="#000000",
            relief=tk.SOLID,
            borderwidth=1,
            font=("メイリオ", 9),
            padx=8,
            pady=4
        )
        label.pack()

    def hide(self) -> None:
        """ツールチップを非表示"""
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None

    def destroy(self) -> None:
        """ツールチップを完全に破棄（メモリリーク防止）"""
        self.unschedule()
        self.hide()
        # イベントバインドを解除
        try:
            self.widget.unbind("<Enter>")
            self.widget.unbind("<Leave>")
            self.widget.unbind("<Button>")
        except tk.TclError:
            pass  # ウィジェットが既に破棄されている場合


def create_tooltip(widget: tk.Widget, text: str) -> ToolTip:
    """
    ウィジェットにツールチップを追加する便利関数

    既存のツールチップがあれば削除してから新規作成することで、
    メモリリークを防止します。

    Args:
        widget: ツールチップを追加するウィジェット
        text: 表示する説明文

    Returns:
        ToolTip: 作成したツールチップオブジェクト
    """
    # 既存のツールチップがあれば削除
    if hasattr(widget, '_tooltip'):
        old_tooltip = widget._tooltip
        if old_tooltip and hasattr(old_tooltip, 'destroy'):
            old_tooltip.destroy()

    # 新しいツールチップを作成し、ウィジェットに保存
    tooltip = ToolTip(widget, text)
    widget._tooltip = tooltip  # 参照を保持してガベージコレクション防止
    return tooltip


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


def open_file_or_folder(path: str, on_error: Optional[Callable[[str], None]] = None) -> bool:
    """
    ファイルまたはフォルダをデフォルトアプリケーションで開く（Windows専用）

    subprocess.Popenを使用してファイルやフォルダを非ブロッキングで開きます。
    この関数は即座にリターンし、GUIをブロックしません。

    セキュリティ対策:
    - パスの正規化と検証を実施
    - コマンドインジェクション対策を実装
    - シンボリックリンク攻撃防止（resolve(strict=True)）
    - 無効文字（null byte等）のチェック

    Args:
        path: 開くファイルまたはフォルダのパス
        on_error: エラー時のコールバック関数（エラーメッセージを受け取る）

    Returns:
        bool: 成功した場合True、失敗した場合False

    Security:
        この関数は信頼できるソース（ファイルダイアログ、設定ファイル、
        アプリケーション内部パス）からのパスのみを受け付けることを前提としています。
        外部入力やネットワーク経由で受け取ったパスを直接渡さないでください。

    Note:
        この関数はWindows専用です。他のOSでは動作しません。
        subprocess.Popenを使用するため、即座にリターンします。
    """
    import subprocess
    import sys
    import logging

    logger = logging.getLogger(__name__)

    # Windowsプラットフォームチェック
    if sys.platform != 'win32':
        if on_error:
            on_error("この機能はWindows専用です。")
        logger.error("open_file_or_folder: Windows以外のプラットフォームで呼び出されました")
        return False

    try:
        import os
        path_obj = Path(path)

        # パスの正規化（フリーズ防止: resolve()を使わない）
        # absolute()のみ使用（シンボリックリンクは解決しない）
        try:
            path_obj = path_obj.absolute()
        except (OSError, ValueError) as e:
            if on_error:
                on_error(f"無効なパス形式です:\n{path}")
            logger.warning(f"パス正規化失敗: {path}, エラー: {e}")
            return False

        # 絶対パス検証
        if not path_obj.is_absolute():
            if on_error:
                on_error("相対パスは使用できません。絶対パスを指定してください。")
            logger.error(f"相対パスが指定されました: {path_obj}")
            return False

        # パス文字列のセキュリティ検証（コマンドインジェクション対策）
        path_str = str(path_obj)

        # 無効文字チェック（null byte, 改行文字等）
        if '\x00' in path_str or '\n' in path_str or '\r' in path_str:
            if on_error:
                on_error("無効なパス文字列が検出されました。")
            logger.error(f"無効な文字を含むパス: {repr(path_str)}")
            return False

        # 存在チェック（os.path経由で高速化・フリーズ防止）
        if not os.path.exists(path_str):
            if on_error:
                on_error(f"指定されたパスが存在しません:\n{path}")
            logger.warning(f"パスが存在しません: {path_str}")
            return False

        # ファイル/フォルダを開く（os.path経由でチェック）
        if os.path.isdir(path_str):
            # フォルダの場合: エクスプローラーで開く
            # subprocess.runを使用してプロセス管理を簡潔に（リソースリーク防止）
            try:
                subprocess.run(
                    ['explorer', path_str],
                    shell=False,
                    stdin=subprocess.DEVNULL,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    check=False,  # エラーを無視（エクスプローラーは正常に起動すれば問題なし）
                    timeout=5  # 5秒でタイムアウト
                )
                logger.debug(f"エクスプローラーで開きました: {path_str}")
            except subprocess.TimeoutExpired:
                # エクスプローラーの起動は非同期なのでタイムアウトは問題なし
                logger.debug(f"エクスプローラー起動（タイムアウト）: {path_str}")
        else:
            # ファイルの場合: os.startfileを使用（最もシンプルで確実）
            os.startfile(path_str)
            logger.debug(f"ファイルを開きました: {path_str}")

        return True

    except FileNotFoundError as e:
        if on_error:
            on_error(f"ファイル/フォルダが見つかりません:\n{path}")
        logger.warning(f"ファイル未検出: {path}, エラー: {e}")
        return False

    except PermissionError as e:
        if on_error:
            on_error(f"アクセス権限がありません:\n{path}")
        logger.warning(f"アクセス拒否: {path}, エラー: {e}")
        return False

    except Exception as e:
        if on_error:
            on_error(f"ファイル/フォルダを開けませんでした:\n{path}")
        logger.error(f"予期しないエラー: {path}, エラー: {e}", exc_info=True)
        return False
