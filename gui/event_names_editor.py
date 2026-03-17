"""
行事名エディタウィジェット

行事名の表示・編集UIを提供する再利用可能なウィジェット
"""
import logging
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from typing import TYPE_CHECKING, Callable, Dict, Optional

from gui.styles import FONTS

if TYPE_CHECKING:
    from infrastructure.config_loader import ConfigLoader

logger = logging.getLogger(__name__)


class EventNamesEditor:
    """
    行事名の表示・編集を担当するウィジェット

    Notebook（タブビュー）を内包し、各カテゴリの行事名リストの
    CRUD操作（追加・編集・削除・並べ替え・デフォルト復元）を提供する。

    Args:
        parent: 親フレーム
        config: ConfigLoaderインスタンス
        status_callback: ステータス更新用コールバック（省略可）
    """

    def __init__(
        self,
        parent: tk.Frame,
        config: "ConfigLoader",
        status_callback: Optional[Callable[[str], None]] = None
    ) -> None:
        self.parent = parent
        self.config = config
        self.status_callback = status_callback
        self.event_listboxes: Dict[str, tk.Listbox] = {}
        self.event_categories: Dict[str, str] = {
            "school_events": "学校行事名 (D列)",
            "student_council_events": "児童会行事名 (C列)",
            "other_activities": "その他の活動 (C列)"
        }
        self._create_ui()

    def _update_status(self, message: str) -> None:
        """
        ステータスバーを更新

        Args:
            message: 表示するメッセージ
        """
        if self.status_callback:
            self.status_callback(message)

    def _create_ui(self) -> None:
        """タブビューとリストボックスパネルを構築"""
        self.event_tabs = ttk.Notebook(self.parent)
        self.event_tabs.pack(fill="both", expand=True, padx=10, pady=5)

        for category, tab_name in self.event_categories.items():
            tab_frame = tk.Frame(self.event_tabs)
            self.event_tabs.add(tab_frame, text=tab_name)
            self._create_event_listbox_panel(tab_frame, category)

    def _create_event_listbox_panel(self, parent: tk.Frame, category: str) -> None:
        """
        リストボックスパネルを作成

        Args:
            parent: 親フレーム
            category: 行事カテゴリキー
        """
        # 左右分割
        container = tk.Frame(parent)
        container.pack(fill="both", expand=True, padx=5, pady=5)

        # 左: リストボックス
        list_frame = tk.Frame(container)
        list_frame.pack(side="left", fill="both", expand=True)

        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")

        listbox = tk.Listbox(
            list_frame, yscrollcommand=scrollbar.set,
            font=FONTS['small'], selectmode="single"
        )
        listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=listbox.yview)

        self.event_listboxes[category] = listbox
        self._load_event_names_to_listbox(category)

        # 右: ボタン
        button_panel = tk.Frame(container)
        button_panel.pack(side="right", fill="y", padx=(8, 0))

        buttons = [
            ("追加", lambda: self._on_add_event_name(category)),
            ("編集", lambda: self._on_edit_event_name(category)),
            ("削除", lambda: self._on_delete_event_name(category)),
            ("▲", lambda: self._on_move_up(category)),
            ("▼", lambda: self._on_move_down(category)),
            ("リセット", lambda: self._on_reset_to_default(category)),
        ]
        for text, cmd in buttons:
            tk.Button(
                button_panel, text=text, command=cmd,
                font=FONTS['tiny'], width=8, cursor="hand2"
            ).pack(pady=1)

    def _load_event_names_to_listbox(self, category: str) -> None:
        """
        行事名をリストボックスに読み込み

        Args:
            category: 行事カテゴリキー
        """
        listbox = self.event_listboxes[category]
        listbox.delete(0, tk.END)

        event_names = self.config.get_event_names(category)
        for name in event_names:
            listbox.insert(tk.END, name)

    def reload_event_names(self) -> None:
        """すべてのカテゴリの行事名をリロード（外部から呼び出し可能）"""
        logger.info("行事名エディタの行事名をリロードしています...")
        for category in self.event_categories.keys():
            self._load_event_names_to_listbox(category)
        logger.info("行事名エディタの行事名をリロードしました")

    def _on_add_event_name(self, category: str) -> None:
        """
        行事名を追加

        Args:
            category: 行事カテゴリキー
        """
        new_name = simpledialog.askstring(
            "行事名を追加",
            "新しい行事名を入力してください:",
            parent=self.parent
        )

        if new_name and new_name.strip():
            new_name = new_name.strip()
            event_names = self.config.get_event_names(category)
            event_names.append(new_name)

            try:
                self.config.save_event_names(category, event_names)
                self._load_event_names_to_listbox(category)
                self._update_status(f"行事名を追加: {new_name}")
            except Exception as e:
                logger.error(f"行事名追加エラー: {e}", exc_info=True)
                messagebox.showerror("追加エラー", f"行事名の追加に失敗しました。\n\n詳細: {e}")

    def _on_edit_event_name(self, category: str) -> None:
        """
        行事名を編集

        Args:
            category: 行事カテゴリキー
        """
        listbox = self.event_listboxes[category]
        selection = listbox.curselection()

        if not selection:
            messagebox.showwarning("未選択", "編集する行事名を選択してください。")
            return

        index = selection[0]
        event_names = self.config.get_event_names(category)
        old_name = event_names[index]

        new_name = simpledialog.askstring(
            "行事名を編集",
            "行事名を編集してください:",
            initialvalue=old_name,
            parent=self.parent
        )

        if new_name and new_name.strip() and new_name.strip() != old_name:
            new_name = new_name.strip()
            event_names[index] = new_name

            try:
                self.config.save_event_names(category, event_names)
                self._load_event_names_to_listbox(category)
                listbox.selection_set(index)  # 編集後も同じ位置を選択
                self._update_status(f"行事名を編集: {old_name} → {new_name}")
            except Exception as e:
                logger.error(f"行事名編集エラー: {e}", exc_info=True)
                messagebox.showerror("編集エラー", f"行事名の編集に失敗しました。\n\n詳細: {e}")

    def _on_delete_event_name(self, category: str) -> None:
        """
        行事名を削除

        Args:
            category: 行事カテゴリキー
        """
        listbox = self.event_listboxes[category]
        selection = listbox.curselection()

        if not selection:
            messagebox.showwarning("未選択", "削除する行事名を選択してください。")
            return

        index = selection[0]
        event_names = self.config.get_event_names(category)
        name = event_names[index]

        # 確認ダイアログ
        result = messagebox.askyesno(
            "削除確認",
            f"「{name}」を削除しますか？",
            parent=self.parent
        )

        if result:
            event_names.pop(index)

            try:
                self.config.save_event_names(category, event_names)
                self._load_event_names_to_listbox(category)
                self._update_status(f"行事名を削除: {name}")
            except Exception as e:
                logger.error(f"行事名削除エラー: {e}", exc_info=True)
                messagebox.showerror("削除エラー", f"行事名の削除に失敗しました。\n\n詳細: {e}")

    def _on_move_up(self, category: str) -> None:
        """
        行事名を上へ移動

        Args:
            category: 行事カテゴリキー
        """
        listbox = self.event_listboxes[category]
        selection = listbox.curselection()

        if not selection:
            messagebox.showwarning("未選択", "移動する行事名を選択してください。")
            return

        index = selection[0]

        if index == 0:
            return

        event_names = self.config.get_event_names(category)
        event_names[index], event_names[index - 1] = event_names[index - 1], event_names[index]

        try:
            self.config.save_event_names(category, event_names)
            self._load_event_names_to_listbox(category)
            listbox.selection_set(index - 1)  # 移動後の位置を選択
            self._update_status(f"行事名を上へ移動: {event_names[index - 1]}")
        except Exception as e:
            logger.error(f"行事名移動エラー: {e}", exc_info=True)
            messagebox.showerror("移動エラー", f"行事名の移動に失敗しました。\n\n詳細: {e}")

    def _on_move_down(self, category: str) -> None:
        """
        行事名を下へ移動

        Args:
            category: 行事カテゴリキー
        """
        listbox = self.event_listboxes[category]
        selection = listbox.curselection()

        if not selection:
            messagebox.showwarning("未選択", "移動する行事名を選択してください。")
            return

        index = selection[0]
        event_names = self.config.get_event_names(category)

        if index == len(event_names) - 1:
            return

        event_names[index], event_names[index + 1] = event_names[index + 1], event_names[index]

        try:
            self.config.save_event_names(category, event_names)
            self._load_event_names_to_listbox(category)
            listbox.selection_set(index + 1)  # 移動後の位置を選択
            self._update_status(f"行事名を下へ移動: {event_names[index + 1]}")
        except Exception as e:
            logger.error(f"行事名移動エラー: {e}", exc_info=True)
            messagebox.showerror("移動エラー", f"行事名の移動に失敗しました。\n\n詳細: {e}")

    def _on_reset_to_default(self, category: str) -> None:
        """
        行事名をデフォルトに戻す

        Args:
            category: 行事カテゴリキー
        """
        # 確認ダイアログ
        result = messagebox.askyesno(
            "デフォルトに戻す",
            "行事名をデフォルト値に戻しますか？\n\n現在の設定は失われます。",
            parent=self.parent
        )

        if not result:
            return

        try:
            was_reset = self.config.reset_event_names(category)
            if was_reset:
                self._load_event_names_to_listbox(category)
                self._update_status("行事名をデフォルトに戻しました")
                messagebox.showinfo("完了", "行事名をデフォルト値に戻しました。")
            else:
                self._update_status("既にデフォルト値です")
        except Exception as e:
            logger.error(f"デフォルト復元エラー: {e}", exc_info=True)
            messagebox.showerror("エラー", f"デフォルト値への復元に失敗しました。\n\n詳細: {e}")
