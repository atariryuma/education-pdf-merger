"""
ファイル管理タブ

ファイル名整理・不要シート削除機能のUIを提供
"""
import tkinter as tk
from tkinter import ttk, messagebox
import threading
from typing import Any

from gui.tabs.base_tab import BaseTab
from gui.utils import create_hover_button


class FileTab(BaseTab):
    """ファイル管理タブ"""

    def __init__(self, notebook: ttk.Notebook, config: Any, status_bar: tk.Label) -> None:
        super().__init__(notebook, config, status_bar)
        self._create_ui()
        self.add_to_notebook("📁 ファイル管理")

    def _create_ui(self) -> None:
        """UIを構築"""
        # 機能セクション用のフレーム
        functions_frame = tk.Frame(self.tab)
        functions_frame.pack(fill="x", padx=20, pady=15)

        # 各機能の説明とボタン
        functions = [
            ("ファイル名整理", "📝 PDCAファイルを01に、その他を連番に自動整理", self._run_rename),
            ("不要シート削除", "🗑️ PDCAファイルから様式解説シートと空白シートを削除", self._run_delete_sheets),
        ]

        for i, (title, desc, command) in enumerate(functions):
            frame = tk.LabelFrame(functions_frame, text=title, font=("メイリオ", 11, "bold"))
            frame.pack(fill="x", pady=(0 if i == 0 else 12))

            content_frame = tk.Frame(frame)
            content_frame.pack(fill="x", pady=12, padx=15)

            tk.Label(content_frame, text=desc, font=("メイリオ", 9)).pack(side="left", anchor="w")

            btn = create_hover_button(
                content_frame,
                text=f"▶ {title}を実行",
                command=command,
                color="warning",
                font=("メイリオ", 10),
                width=20,
                height=1
            )
            btn.pack(side="right", padx=10)

        # ログ表示
        self.create_log_frame(height=10)
        self.log("準備完了。実行したい機能のボタンをクリックしてください。", "info")

    def _run_rename(self) -> None:
        """ファイル名整理を実行"""
        def task():
            try:
                self.update_status("ファイル名整理を実行中...")
                self.log("=== ファイル名整理開始 ===", "info")

                import rename_file
                rename_file.main()

                self.log("=== ファイル名整理完了 ===", "success")
                self.update_status("ファイル名整理が完了しました")
                messagebox.showinfo("✅ 完了", "ファイル名整理が完了しました！")
            except Exception as e:
                self.log(f"エラー: {e}", "error")
                self.update_status("ファイル名整理でエラーが発生しました")
                messagebox.showerror("❌ 実行エラー", f"エラーが発生しました。\n\n詳細:\n{e}")

        thread = threading.Thread(target=task, daemon=True)
        thread.start()

    def _run_delete_sheets(self) -> None:
        """不要シート削除を実行"""
        def task():
            try:
                self.update_status("不要シート削除を実行中...")
                self.log("=== 不要シート削除開始 ===", "info")

                import delete
                delete.main()

                self.log("=== 不要シート削除完了 ===", "success")
                self.update_status("不要シート削除が完了しました")
                messagebox.showinfo("✅ 完了", "不要シート削除が完了しました！")
            except Exception as e:
                self.log(f"エラー: {e}", "error")
                self.update_status("不要シート削除でエラーが発生しました")
                messagebox.showerror("❌ 実行エラー", f"エラーが発生しました。\n\n詳細:\n{e}")

        thread = threading.Thread(target=task, daemon=True)
        thread.start()
