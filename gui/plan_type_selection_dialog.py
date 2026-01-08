"""
計画種別選択ダイアログモジュール

フォルダ構造判定が曖昧な場合にユーザーに選択を促す
"""
import tkinter as tk
from tkinter import ttk
from typing import Callable


class PlanTypeSelectionDialog(tk.Toplevel):
    """計画種別選択ダイアログ"""

    def __init__(
        self,
        parent: tk.Widget,
        detection_result,
        callback: Callable[[str], None]
    ):
        """
        Args:
            parent: 親ウィジェット
            detection_result: フォルダ構造の判定結果（DetectionResult）
            callback: 選択結果のコールバック(plan_type: str) -> None
        """
        super().__init__(parent)
        self.detection_result = detection_result
        self.callback = callback

        # ウィンドウ設定
        self.title("計画種別の選択")
        self.geometry("550x450")
        self.resizable(False, False)

        # モーダルダイアログ
        self.transient(parent)
        self.grab_set()

        # 中央配置
        self._center_window()

        # UI構築
        self._create_widgets()

    def _center_window(self):
        """ウィンドウを画面中央に配置"""
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - (550 // 2)
        y = (self.winfo_screenheight() // 2) - (450 // 2)
        self.geometry(f"+{x}+{y}")

    def _create_widgets(self):
        """ウィジェット作成"""
        main_frame = tk.Frame(self, padx=30, pady=30)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # タイトル
        title_label = tk.Label(
            main_frame,
            text="計画種別を選択してください",
            font=("Yu Gothic UI", 14, "bold"),
            fg="#333333"
        )
        title_label.pack(pady=(0, 10))

        # メッセージ
        message_label = tk.Label(
            main_frame,
            text="フォルダ構造から自動判定できませんでした。\n計画種別を手動で選択してください。",
            font=("Yu Gothic UI", 10),
            justify=tk.CENTER,
            fg="#666666"
        )
        message_label.pack(pady=(0, 20))

        # 判定情報フレーム
        info_frame = tk.LabelFrame(
            main_frame,
            text="フォルダ構造の分析結果",
            font=("Yu Gothic UI", 10, "bold"),
            padx=15,
            pady=15
        )
        info_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))

        evidence = self.detection_result.evidence

        info_items = [
            ("メインディレクトリ数", f"{evidence.get('main_dir_count', 0)}個"),
            ("ルート直下のファイル数", f"{evidence.get('root_file_count', 0)}個"),
            ("最大階層深度", f"{evidence.get('max_depth', 0)}層"),
            ("ルートファイル比率", f"{evidence.get('root_file_ratio', 0):.1%}"),
            ("教育計画スコア", f"{evidence.get('education_score', 0):.1f}"),
            ("行事計画スコア", f"{evidence.get('event_score', 0):.1f}")
        ]

        for i, (label_text, value_text) in enumerate(info_items):
            label = tk.Label(
                info_frame,
                text=f"{label_text}:",
                font=("Yu Gothic UI", 10),
                anchor="w"
            )
            label.grid(row=i, column=0, sticky="w", pady=5)

            value = tk.Label(
                info_frame,
                text=value_text,
                font=("Yu Gothic UI", 10, "bold"),
                anchor="e"
            )
            value.grid(row=i, column=1, sticky="e", pady=5, padx=(20, 0))

        info_frame.columnconfigure(0, weight=1)
        info_frame.columnconfigure(1, weight=1)

        # 選択ボタンフレーム
        button_frame = tk.Frame(main_frame)
        button_frame.pack(pady=(0, 10))

        # 教育計画ボタン
        education_btn = tk.Button(
            button_frame,
            text="📚 教育計画\n（3層構造）",
            command=lambda: self._on_select("education"),
            bg="#4CAF50",
            fg="white",
            font=("Yu Gothic UI", 11, "bold"),
            width=15,
            height=3,
            cursor="hand2"
        )
        education_btn.pack(side=tk.LEFT, padx=10)

        # 行事計画ボタン
        event_btn = tk.Button(
            button_frame,
            text="📅 行事計画\n（2層構造）",
            command=lambda: self._on_select("event"),
            bg="#2196F3",
            fg="white",
            font=("Yu Gothic UI", 11, "bold"),
            width=15,
            height=3,
            cursor="hand2"
        )
        event_btn.pack(side=tk.LEFT, padx=10)

        # キャンセルボタン
        cancel_btn = tk.Button(
            main_frame,
            text="キャンセル",
            command=self._on_cancel,
            font=("Yu Gothic UI", 10),
            cursor="hand2"
        )
        cancel_btn.pack()

    def _on_select(self, plan_type: str):
        """選択ボタンクリック時の処理"""
        self.destroy()
        if self.callback:
            self.callback(plan_type)

    def _on_cancel(self):
        """キャンセルボタンクリック時の処理"""
        self.destroy()
