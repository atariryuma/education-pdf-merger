"""
ファイル管理タブ

ファイル名整理・不要シート削除機能のUIを提供
"""
import tkinter as tk
from tkinter import ttk
from typing import TYPE_CHECKING

from gui.tabs.base_tab import BaseTab

if TYPE_CHECKING:
    from config_loader import ConfigLoader


class FileTab(BaseTab):
    """ファイル管理タブ"""

    def __init__(self, notebook: ttk.Notebook, config: "ConfigLoader", status_bar: tk.Label) -> None:
        super().__init__(notebook, config, status_bar)
        self._create_ui()
        self.add_to_notebook("📁 ファイル管理")

    def _create_ui(self) -> None:
        """UIを構築"""
        # 情報フレーム
        info_frame = tk.LabelFrame(self.tab, text="📋 ファイル管理機能", font=("メイリオ", 11, "bold"))
        info_frame.pack(fill="x", padx=20, pady=15)

        info_text = "このタブでは、将来的にファイル名整理や不要シート削除などの機能が提供される予定です。"
        tk.Label(info_frame, text=info_text, justify="left", font=("メイリオ", 10)).pack(pady=15, padx=15)

        # ログ表示
        self.create_log_frame(height=10)
        self.log("ファイル管理機能は現在開発中です。", "info")
