"""
一太郎PDF変換中の警告ダイアログモジュール

一太郎変換中にキーボード入力を避けるよう、ユーザーに警告を表示する
非モーダルダイアログを提供する。
"""
import tkinter as tk
from tkinter import ttk
from typing import Callable, Optional


class IchitaroConversionDialog(tk.Toplevel):
    """一太郎PDF変換中の警告ダイアログ（非モーダル、常に最前面）"""

    def __init__(self, parent: tk.Widget, cancel_callback: Optional[Callable] = None):
        """
        Args:
            parent: 親ウィジェット
            cancel_callback: キャンセルボタンのコールバック関数
        """
        super().__init__(parent)
        self.cancel_callback = cancel_callback

        # ウィンドウ設定
        self.title("一太郎PDF変換中")
        self.geometry("500x200")
        self.resizable(False, False)

        # 常に最前面（非モーダル）
        self.attributes('-topmost', True)

        # 親ウィンドウとの関連付けのみ（grab_set()は使わない）
        self.transient(parent)

        # 中央配置
        self._center_window()

        # UI構築
        self._create_widgets()

    def _center_window(self):
        """ウィンドウを画面中央に配置"""
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (500 // 2)
        y = (self.winfo_screenheight() // 2) - (200 // 2)
        self.geometry(f"+{x}+{y}")

    def _create_widgets(self):
        """ウィジェット作成"""
        # メインフレーム
        main_frame = tk.Frame(self, padx=30, pady=30)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # アイコン + タイトル
        title_label = tk.Label(
            main_frame,
            text="⚠️ 一太郎PDF変換中",
            font=("Yu Gothic UI", 16, "bold"),
            fg="#ff6b00"
        )
        title_label.pack(pady=(0, 20))

        # メッセージ
        self.message_label = tk.Label(
            main_frame,
            text="一太郎ファイルをPDFに変換しています。\nしばらくお待ちください...",
            font=("Yu Gothic UI", 11),
            justify=tk.CENTER
        )
        self.message_label.pack(pady=(0, 20))

        # プログレスバー
        self.progress = ttk.Progressbar(
            main_frame,
            mode="indeterminate",
            length=400
        )
        self.progress.pack(pady=(0, 20))
        self.progress.start(10)

        # 注意書き
        note_label = tk.Label(
            main_frame,
            text="※ 変換中はキーボード操作を避けてください",
            font=("Yu Gothic UI", 9),
            fg="#666666"
        )
        note_label.pack()

        # キャンセルボタン
        if self.cancel_callback:
            cancel_btn = tk.Button(
                main_frame,
                text="キャンセル",
                command=self._on_cancel,
                bg="#dc3545",
                fg="white",
                font=("Yu Gothic UI", 10),
                cursor="hand2",
                padx=20,
                pady=5
            )
            cancel_btn.pack(pady=(10, 0))

    def update_message(self, message: str):
        """
        メッセージを更新

        Args:
            message: 新しいメッセージ
        """
        self.message_label.config(text=message)
        self.update_idletasks()

    def _on_cancel(self):
        """キャンセルボタンクリック時の処理"""
        if self.cancel_callback:
            self.cancel_callback()
        self.close()

    def close(self):
        """ダイアログを閉じる"""
        try:
            self.progress.stop()
            # grab_release()は不要（grab_set()していないため）
            self.destroy()
        except:
            pass
