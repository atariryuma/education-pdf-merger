"""
設定タブ

アプリケーション設定のUIを提供
"""
import logging
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Any, Callable, Optional
from pathlib import Path

from gui.tabs.base_tab import BaseTab
from gui.utils import create_hover_button, open_file_or_folder, thread_safe_call
from path_validator import PathValidator
from year_utils import calculate_year_short

logger = logging.getLogger(__name__)


class SettingsTab(BaseTab):
    """設定タブ"""

    def __init__(
        self,
        notebook: ttk.Notebook,
        config: Any,
        status_bar: tk.Label,
        year_var: tk.StringVar,
        year_short_var: tk.StringVar,
        gdrive_var: tk.StringVar,
        temp_var: tk.StringVar,
        gs_var: tk.StringVar,
        on_reload: Callable[[], None]
    ) -> None:
        super().__init__(notebook, config, status_bar)
        self.year_var = year_var
        self.year_short_var = year_short_var
        self.gdrive_var = gdrive_var
        self.temp_var = temp_var
        self.gs_var = gs_var
        self.on_reload = on_reload

        # 年度変更時に自動でyear_shortを更新
        self.year_var.trace_add('write', self._on_year_changed)

        self._create_ui()
        self.add_to_notebook("⚙️ 設定")

    def _on_year_changed(self, *args) -> None:
        """年度が変更されたときに和暦を自動更新"""
        year = self.year_var.get()
        if year.isdigit() and len(year) == 4:
            year_short = calculate_year_short(year)
            self.year_short_var.set(year_short)

    def _show_file_open_error(self, error_msg: str) -> None:
        """
        ファイル/フォルダを開く際のエラーを表示（共通処理）

        Args:
            error_msg: エラーメッセージ
        """
        messagebox.showerror("エラー", error_msg)

    def _create_ui(self) -> None:
        """UIを構築"""
        # スクロール可能なメインコンテナ（BaseTabの共通メソッドを使用）
        self.canvas, _scrollbar, self.scrollable_frame = self.create_scrollable_container()

        # メインコンテナ（スクロール可能フレーム内）
        main_container = self.scrollable_frame

        # 説明フレーム（初心者向け）
        help_frame = tk.LabelFrame(main_container, text="💡 設定について", font=("メイリオ", 10, "bold"))
        help_frame.pack(fill="x", pady=(0, 10))

        help_text = (
            "このタブでは、アプリケーションの基本設定を行います。\n\n"
            "📁 = フォルダを選択　│　📂 = フォルダを開く\n"
            "📄 = ファイルを選択　│　🔍 = 自動検索\n\n"
            "⚠️ 設定を変更したら、必ず「💾 保存」ボタンをクリックしてください。"
        )
        tk.Label(
            help_frame,
            text=help_text,
            justify="left",
            font=("メイリオ", 9),
            fg="#333",
            padx=15,
            pady=10
        ).pack(anchor="w")

        # 共通のラベル幅とパディング
        LABEL_WIDTH = 16
        PAD_Y = 5

        # --- 年度情報 ---
        year_frame = tk.LabelFrame(main_container, text="📅 年度情報", font=("メイリオ", 10, "bold"))
        year_frame.pack(fill="x", pady=(0, 8))

        tk.Label(year_frame, text="年度（西暦）:", width=LABEL_WIDTH, anchor="e").grid(row=0, column=0, sticky="e", padx=(10, 3), pady=PAD_Y)
        tk.Entry(year_frame, textvariable=self.year_var, width=15).grid(row=0, column=1, sticky="w", padx=3, pady=PAD_Y)
        tk.Label(year_frame, text="→", font=("メイリオ", 10)).grid(row=0, column=2, sticky="w", padx=3, pady=PAD_Y)
        tk.Label(year_frame, textvariable=self.year_short_var, font=("メイリオ", 10, "bold"), fg="#1976D2").grid(row=0, column=3, sticky="w", padx=3, pady=PAD_Y)
        tk.Label(year_frame, text="💡 和暦は自動計算", font=("メイリオ", 8), fg="gray").grid(row=1, column=1, columnspan=3, sticky="w", padx=3, pady=(0, 5))

        # --- パス設定 ---
        path_frame = tk.LabelFrame(main_container, text="📂 パス設定", font=("メイリオ", 10, "bold"))
        path_frame.pack(fill="x", pady=8)

        tk.Label(path_frame, text="Google Drive:", width=LABEL_WIDTH, anchor="e").grid(row=0, column=0, sticky="e", padx=(10, 3), pady=PAD_Y)
        tk.Entry(path_frame, textvariable=self.gdrive_var).grid(row=0, column=1, sticky="ew", padx=3, pady=PAD_Y)

        gdrive_btn_frame = tk.Frame(path_frame)
        gdrive_btn_frame.grid(row=0, column=2, padx=(3, 10), pady=PAD_Y)
        tk.Button(gdrive_btn_frame, text="📁", command=lambda: self._browse_folder(self.gdrive_var), width=3).pack(side="left", padx=1)
        tk.Button(gdrive_btn_frame, text="📂", command=lambda: self._open_folder(self.gdrive_var), width=3).pack(side="left", padx=1)

        tk.Label(path_frame, text="一時フォルダ:", width=LABEL_WIDTH, anchor="e").grid(row=1, column=0, sticky="e", padx=(10, 3), pady=PAD_Y)
        tk.Entry(path_frame, textvariable=self.temp_var).grid(row=1, column=1, sticky="ew", padx=3, pady=PAD_Y)

        temp_btn_frame = tk.Frame(path_frame)
        temp_btn_frame.grid(row=1, column=2, padx=(3, 10), pady=PAD_Y)
        tk.Button(temp_btn_frame, text="📁", command=lambda: self._browse_folder(self.temp_var), width=3).pack(side="left", padx=1)
        tk.Button(temp_btn_frame, text="📂", command=self._open_temp_folder, width=3).pack(side="left", padx=1)

        path_frame.columnconfigure(1, weight=1)

        # --- ツール設定 ---
        tool_frame = tk.LabelFrame(main_container, text="🔧 ツール設定", font=("メイリオ", 10, "bold"))
        tool_frame.pack(fill="x", pady=8)

        tk.Label(tool_frame, text="Ghostscript:", width=LABEL_WIDTH, anchor="e").grid(row=0, column=0, sticky="e", padx=(10, 3), pady=PAD_Y)
        tk.Entry(tool_frame, textvariable=self.gs_var).grid(row=0, column=1, sticky="ew", padx=3, pady=PAD_Y)

        gs_btn_frame = tk.Frame(tool_frame)
        gs_btn_frame.grid(row=0, column=2, padx=(3, 10), pady=PAD_Y)
        tk.Button(gs_btn_frame, text="📄", command=self._browse_gs_file, width=3).pack(side="left", padx=1)
        tk.Button(gs_btn_frame, text="🔍 自動検出", command=self._auto_detect_ghostscript, font=("メイリオ", 8)).pack(side="left", padx=1)

        # Ghostscriptステータス表示
        self.gs_status_label = tk.Label(tool_frame, text="", fg="gray", font=("メイリオ", 8))
        self.gs_status_label.grid(row=1, column=1, columnspan=2, sticky="w", padx=3, pady=(0, 3))
        self._update_gs_status()

        tool_frame.columnconfigure(1, weight=1)

        # --- 一太郎設定 ---
        ichitaro_frame = tk.LabelFrame(main_container, text="📝 一太郎変換設定", font=("メイリオ", 10, "bold"))
        ichitaro_frame.pack(fill="x", pady=8)

        # 設定値の読み込み
        self.max_retries_var = tk.StringVar(value=str(self.config.get('ichitaro', 'max_retries') or 3))
        self.save_wait_var = tk.StringVar(value=str(self.config.get('ichitaro', 'save_wait_seconds') or 20))

        # 設定行: リトライ回数、保存待機時間、テストボタン
        settings_row1 = tk.Frame(ichitaro_frame)
        settings_row1.pack(fill="x", padx=10, pady=PAD_Y)
        tk.Label(settings_row1, text="リトライ:").pack(side="left")
        tk.Entry(settings_row1, textvariable=self.max_retries_var, width=3).pack(side="left", padx=(3, 0))
        tk.Label(settings_row1, text="回").pack(side="left", padx=(2, 15))
        tk.Label(settings_row1, text="保存待機:").pack(side="left")
        tk.Entry(settings_row1, textvariable=self.save_wait_var, width=3).pack(side="left", padx=(3, 0))
        tk.Label(settings_row1, text="秒").pack(side="left", padx=(2, 15))
        tk.Button(settings_row1, text="🧪 テスト", command=self._test_ichitaro_conversion, font=("メイリオ", 8)).pack(side="left", padx=5)

        # 説明ラベル
        help_label = tk.Label(
            ichitaro_frame,
            text="💡 Microsoft Print to PDFを自動選択します（環境非依存）",
            fg="#0066cc",
            font=("メイリオ", 8)
        )
        help_label.pack(anchor="w", padx=10, pady=(0, 3))

        # ステータス表示
        self.ichitaro_status_label = tk.Label(
            ichitaro_frame,
            text="処理手順: Ctrl+P → プリンター自動選択 → Enter → ファイル名 → Enter",
            fg="#666",
            font=("メイリオ", 8)
        )
        self.ichitaro_status_label.pack(anchor="w", padx=10, pady=(0, 3))

        # ログファイルボタン
        log_button_frame = tk.Frame(ichitaro_frame)
        log_button_frame.pack(anchor="w", padx=10, pady=(5, 3))
        tk.Button(log_button_frame, text="📄 ログファイルを開く", command=self._open_log_file, font=("メイリオ", 8)).pack(side="left")

        # --- 行事名設定（折りたたみ式） ---
        event_names_container = tk.Frame(main_container)
        event_names_container.pack(fill="x", pady=8)

        # トグルボタン付きヘッダー
        event_header_frame = tk.Frame(event_names_container)
        event_header_frame.pack(fill="x")

        self.event_names_expanded = tk.BooleanVar(value=False)  # デフォルトで折りたたみ

        self.event_toggle_button = tk.Button(
            event_header_frame,
            text="▶ 行事名設定（Excel転記用）を展開",
            command=self._toggle_event_names_section,
            font=("メイリオ", 10, "bold"),
            relief="flat",
            anchor="w",
            cursor="hand2",
            bg="#f0f0f0"
        )
        self.event_toggle_button.pack(fill="x", padx=5, pady=2)

        # 折りたたみ可能なコンテンツフレーム
        self.event_names_content = tk.Frame(event_names_container)
        # デフォルトでは非表示（pack_forget状態）

        # タブビュー作成
        self.event_tabs = ttk.Notebook(self.event_names_content)
        self.event_tabs.pack(fill="both", expand=True, padx=10, pady=5)

        # 各カテゴリのタブを作成
        self.event_listboxes = {}
        self.event_categories = {
            "school_events": "学校行事名 (D列)",
            "student_council_events": "児童会行事名 (C列)",
            "other_activities": "その他の活動 (C列)"
        }

        for category, tab_name in self.event_categories.items():
            tab_frame = tk.Frame(self.event_tabs)
            self.event_tabs.add(tab_frame, text=tab_name)
            self._create_event_listbox_panel(tab_frame, category)

        # 説明ラベル（折りたたみ時も表示）
        tk.Label(
            event_names_container,
            text="💡 Excelタブから行事名を読み込めます。カスタマイズする場合は上記を展開してください。",
            font=("メイリオ", 8),
            fg="#666"
        ).pack(anchor="w", padx=15, pady=(3, 0))

        # --- ボタン行 ---
        button_frame = tk.Frame(main_container)
        button_frame.pack(pady=15)

        save_btn = create_hover_button(
            button_frame,
            text="💾 保存 (Ctrl+S)",
            command=self.save_settings,
            color="primary",
            font=("メイリオ", 9, "bold"),
            width=14,
            height=1
        )
        save_btn.pack(side="left", padx=5)

        reload_btn = tk.Button(
            button_frame,
            text="🔄 再読み込み (Ctrl+R)",
            command=self.reload_settings,
            font=("メイリオ", 9),
            width=18,
            height=1,
            cursor="hand2"
        )
        reload_btn.pack(side="left", padx=5)

        edit_btn = tk.Button(
            button_frame,
            text="📝 config.json編集",
            command=self.open_config_file,
            font=("メイリオ", 9),
            width=16,
            height=1,
            cursor="hand2"
        )
        edit_btn.pack(side="left", padx=5)

    def _toggle_event_names_section(self) -> None:
        """行事名設定セクションを展開/折りたたみ"""
        if self.event_names_expanded.get():
            # 折りたたむ
            self.event_names_content.pack_forget()
            self.event_toggle_button.config(text="▶ 行事名設定（Excel転記用）を展開")
            self.event_names_expanded.set(False)
        else:
            # 展開
            self.event_names_content.pack(fill="both", expand=True, padx=5, pady=5)
            self.event_toggle_button.config(text="▼ 行事名設定（Excel転記用）を折りたたむ")
            self.event_names_expanded.set(True)

    def _create_event_listbox_panel(self, parent: tk.Frame, category: str) -> None:
        """リストボックスパネルを作成"""
        # メインコンテナ（左右分割）
        container = tk.Frame(parent)
        container.pack(fill="both", expand=True, padx=5, pady=5)

        # 左側: リストボックス
        list_frame = tk.Frame(container)
        list_frame.pack(side="left", fill="both", expand=True)

        # スクロールバー付きリストボックス
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")

        listbox = tk.Listbox(
            list_frame,
            yscrollcommand=scrollbar.set,
            font=("メイリオ", 9),
            height=12,
            selectmode="single"
        )
        listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=listbox.yview)

        # リストボックスを保存
        self.event_listboxes[category] = listbox

        # 行事名をロード
        self._load_event_names_to_listbox(category)

        # 右側: ボタンパネル
        button_panel = tk.Frame(container)
        button_panel.pack(side="right", fill="y", padx=(10, 0))

        # ボタン作成
        tk.Button(
            button_panel,
            text="➕ 追加",
            command=lambda: self._on_add_event_name(category),
            font=("メイリオ", 9),
            width=12,
            cursor="hand2"
        ).pack(pady=3)

        tk.Button(
            button_panel,
            text="✏️ 編集",
            command=lambda: self._on_edit_event_name(category),
            font=("メイリオ", 9),
            width=12,
            cursor="hand2"
        ).pack(pady=3)

        tk.Button(
            button_panel,
            text="🗑️ 削除",
            command=lambda: self._on_delete_event_name(category),
            font=("メイリオ", 9),
            width=12,
            cursor="hand2"
        ).pack(pady=3)

        tk.Label(button_panel, text="").pack(pady=3)  # スペーサー

        tk.Button(
            button_panel,
            text="⬆️ 上へ",
            command=lambda: self._on_move_up(category),
            font=("メイリオ", 9),
            width=12,
            cursor="hand2"
        ).pack(pady=3)

        tk.Button(
            button_panel,
            text="⬇️ 下へ",
            command=lambda: self._on_move_down(category),
            font=("メイリオ", 9),
            width=12,
            cursor="hand2"
        ).pack(pady=3)

        tk.Label(button_panel, text="").pack(pady=8)  # スペーサー

        tk.Button(
            button_panel,
            text="🔄 デフォルトに戻す",
            command=lambda: self._on_reset_to_default(category),
            font=("メイリオ", 8),
            width=12,
            cursor="hand2",
            fg="blue"
        ).pack(pady=3)

    def _load_event_names_to_listbox(self, category: str) -> None:
        """行事名をリストボックスに読み込み"""
        listbox = self.event_listboxes[category]
        listbox.delete(0, tk.END)

        event_names = self.config.get_event_names(category)
        for name in event_names:
            listbox.insert(tk.END, name)

    def reload_event_names(self) -> None:
        """すべてのカテゴリの行事名をリロード（外部から呼び出し可能）"""
        logger.info("設定タブの行事名をリロードしています...")
        for category in self.event_categories.keys():
            self._load_event_names_to_listbox(category)
        logger.info("設定タブの行事名をリロードしました")

    def _on_add_event_name(self, category: str) -> None:
        """行事名を追加"""
        from tkinter import simpledialog

        new_name = simpledialog.askstring(
            "行事名を追加",
            "新しい行事名を入力してください:",
            parent=self.tab
        )

        if new_name and new_name.strip():
            new_name = new_name.strip()
            event_names = self.config.get_event_names(category)
            event_names.append(new_name)

            try:
                self.config.save_event_names(category, event_names)
                self._load_event_names_to_listbox(category)
                self.update_status(f"行事名を追加: {new_name}")
            except Exception as e:
                logger.error(f"行事名追加エラー: {e}", exc_info=True)
                messagebox.showerror("追加エラー", f"行事名の追加に失敗しました。\n\n詳細: {e}")

    def _on_edit_event_name(self, category: str) -> None:
        """行事名を編集"""
        from tkinter import simpledialog

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
            parent=self.tab
        )

        if new_name and new_name.strip() and new_name.strip() != old_name:
            new_name = new_name.strip()
            event_names[index] = new_name

            try:
                self.config.save_event_names(category, event_names)
                self._load_event_names_to_listbox(category)
                listbox.selection_set(index)  # 編集後も同じ位置を選択
                self.update_status(f"行事名を編集: {old_name} → {new_name}")
            except Exception as e:
                logger.error(f"行事名編集エラー: {e}", exc_info=True)
                messagebox.showerror("編集エラー", f"行事名の編集に失敗しました。\n\n詳細: {e}")

    def _on_delete_event_name(self, category: str) -> None:
        """行事名を削除"""
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
            parent=self.tab
        )

        if result:
            event_names.pop(index)

            try:
                self.config.save_event_names(category, event_names)
                self._load_event_names_to_listbox(category)
                self.update_status(f"行事名を削除: {name}")
            except Exception as e:
                logger.error(f"行事名削除エラー: {e}", exc_info=True)
                messagebox.showerror("削除エラー", f"行事名の削除に失敗しました。\n\n詳細: {e}")

    def _on_move_up(self, category: str) -> None:
        """行事名を上へ移動"""
        listbox = self.event_listboxes[category]
        selection = listbox.curselection()

        if not selection:
            messagebox.showwarning("未選択", "移動する行事名を選択してください。")
            return

        index = selection[0]

        if index == 0:
            messagebox.showinfo("移動不可", "既に最上位です。")
            return

        event_names = self.config.get_event_names(category)
        event_names[index], event_names[index - 1] = event_names[index - 1], event_names[index]

        try:
            self.config.save_event_names(category, event_names)
            self._load_event_names_to_listbox(category)
            listbox.selection_set(index - 1)  # 移動後の位置を選択
            self.update_status(f"行事名を上へ移動: {event_names[index - 1]}")
        except Exception as e:
            logger.error(f"行事名移動エラー: {e}", exc_info=True)
            messagebox.showerror("移動エラー", f"行事名の移動に失敗しました。\n\n詳細: {e}")

    def _on_move_down(self, category: str) -> None:
        """行事名を下へ移動"""
        listbox = self.event_listboxes[category]
        selection = listbox.curselection()

        if not selection:
            messagebox.showwarning("未選択", "移動する行事名を選択してください。")
            return

        index = selection[0]
        event_names = self.config.get_event_names(category)

        if index == len(event_names) - 1:
            messagebox.showinfo("移動不可", "既に最下位です。")
            return

        event_names[index], event_names[index + 1] = event_names[index + 1], event_names[index]

        try:
            self.config.save_event_names(category, event_names)
            self._load_event_names_to_listbox(category)
            listbox.selection_set(index + 1)  # 移動後の位置を選択
            self.update_status(f"行事名を下へ移動: {event_names[index + 1]}")
        except Exception as e:
            logger.error(f"行事名移動エラー: {e}", exc_info=True)
            messagebox.showerror("移動エラー", f"行事名の移動に失敗しました。\n\n詳細: {e}")

    def _on_reset_to_default(self, category: str) -> None:
        """行事名をデフォルトに戻す"""
        # 確認ダイアログ
        result = messagebox.askyesno(
            "デフォルトに戻す",
            "行事名をデフォルト値に戻しますか？\n\n現在の設定は失われます。",
            parent=self.tab
        )

        if not result:
            return

        try:
            was_reset = self.config.reset_event_names(category)
            if was_reset:
                self._load_event_names_to_listbox(category)
                self.update_status("行事名をデフォルトに戻しました")
                messagebox.showinfo("完了", "行事名をデフォルト値に戻しました。")
            else:
                messagebox.showinfo("完了", "既にデフォルト値です。")
        except Exception as e:
            logger.error(f"デフォルト復元エラー: {e}", exc_info=True)
            messagebox.showerror("エラー", f"デフォルト値への復元に失敗しました。\n\n詳細: {e}")

    def _browse_folder(self, var: tk.StringVar) -> None:
        """フォルダを参照（PathValidatorベース）"""
        try:
            current_path_str = var.get().strip()
            initial_dir = PathValidator.get_safe_initial_dir(current_path_str, Path.home())

            directory = filedialog.askdirectory(title="フォルダを選択", initialdir=str(initial_dir))
            if directory:
                is_valid, error_msg, validated_path = PathValidator.validate_directory(
                    directory, must_exist=True
                )
                if is_valid and validated_path:
                    var.set(str(validated_path))
                    self.update_status(f"フォルダを選択: {validated_path.name}")
                else:
                    messagebox.showerror("パスエラー", error_msg or "フォルダが無効です")
        except Exception as e:
            messagebox.showerror("参照エラー", f"フォルダの参照中にエラーが発生しました。\n\n詳細: {e}")

    def _browse_gs_file(self) -> None:
        """Ghostscript実行ファイルを参照（フリーズ防止版）"""
        try:
            current_path = self.gs_var.get().strip()
            # ローカルパス（C:ドライブ）のみチェック（フリーズ防止）
            if current_path:
                # ネットワークパスかチェック
                if not current_path.startswith('\\\\') and len(current_path) >= 3 and current_path[1] == ':':
                    drive = current_path[0].upper()
                    if drive in ['C', 'D', 'E'] and os.path.exists(current_path) and os.path.isfile(current_path):
                        initial_dir = os.path.dirname(current_path)
                    else:
                        initial_dir = "C:\\Program Files"
                else:
                    initial_dir = "C:\\Program Files"
            elif os.path.exists("C:\\Program Files\\gs"):
                initial_dir = "C:\\Program Files\\gs"
            else:
                initial_dir = "C:\\Program Files"

            file_path = filedialog.askopenfilename(
                title="Ghostscript実行ファイルを選択",
                initialdir=initial_dir,
                filetypes=[("実行ファイル", "*.exe"), ("すべて", "*.*")]
            )
            if file_path:
                # PathValidatorで検証
                is_valid, error_msg, validated_path = PathValidator.validate_file_path(
                    file_path,
                    must_exist=True
                )
                if not is_valid:
                    messagebox.showerror("パス検証エラー", error_msg)
                    return

                self.gs_var.set(str(validated_path))
                self._update_gs_status_sync()
                self.update_status(f"Ghostscript: {validated_path.name}")
        except Exception as e:
            messagebox.showerror("参照エラー", f"ファイルの参照中にエラーが発生しました。\n\n詳細: {e}")

    def _open_folder(self, var: tk.StringVar) -> None:
        """フォルダをエクスプローラーで開く"""
        folder_path_str = var.get().strip()

        if not folder_path_str:
            messagebox.showwarning("警告", "フォルダパスが設定されていません。")
            return

        if open_file_or_folder(folder_path_str, self._show_file_open_error):
            self.update_status(f"フォルダを開きました: {Path(folder_path_str).name}")

    def _open_temp_folder(self) -> None:
        """一時フォルダをエクスプローラーで開く（フリーズ防止版）"""
        temp_path_str = self.temp_var.get().strip()

        # パスが空の場合はデフォルトパスを使用
        if not temp_path_str:
            appdata = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
            temp_path_str = os.path.join(appdata, 'PDFMergeSystem', 'temp')

        temp_path = Path(temp_path_str)

        # フォルダが存在しない場合は作成（os.path経由でフリーズ防止）
        temp_path_str_final = str(temp_path)
        if not os.path.exists(temp_path_str_final):
            try:
                temp_path.mkdir(parents=True, exist_ok=True)
                self.update_status(f"一時フォルダを作成しました: {temp_path.name}")
            except Exception as e:
                messagebox.showerror("エラー", f"一時フォルダの作成に失敗しました。\n\n{e}")
                return

        # エクスプローラーで開く（非同期）
        if open_file_or_folder(str(temp_path), self._show_file_open_error):
            self.update_status("一時フォルダを開きました")

    def save_settings(self) -> None:
        """設定を保存（入力検証付き - ベストプラクティス準拠）"""
        year = self.year_var.get().strip()

        if not year:
            messagebox.showerror("入力エラー", "年度情報は必須です。")
            return

        # year_shortは自動計算（update_yearに渡さない）
        self.config.update_year(year)
        self.config.set('base_paths', 'google_drive', value=self.gdrive_var.get())
        self.config.set('base_paths', 'local_temp', value=self.temp_var.get())
        self.config.set('ghostscript', 'executable', value=self.gs_var.get())

        # 一太郎設定の保存（入力検証付き）
        validation_errors = []

        try:
            retry_value = int(self.max_retries_var.get())
            if retry_value < 0 or retry_value > 10:
                validation_errors.append("• リトライ回数は0～10の範囲で入力してください")
            else:
                self.config.set('ichitaro', 'max_retries', value=retry_value)
        except ValueError:
            validation_errors.append("• リトライ回数は整数で入力してください")

        try:
            wait_value = int(self.save_wait_var.get())
            if wait_value < 5 or wait_value > 120:
                validation_errors.append("• 保存待機時間は5～120秒の範囲で入力してください")
            else:
                self.config.set('ichitaro', 'save_wait_seconds', value=wait_value)
        except ValueError:
            validation_errors.append("• 保存待機時間は整数で入力してください")

        # 検証エラーがあれば表示して保存を中断
        if validation_errors:
            error_message = "入力値に誤りがあります:\n\n" + "\n".join(validation_errors)
            messagebox.showwarning("入力エラー", error_message)
            return

        try:
            self.config.save_config()
            self.update_status("設定を保存しました")
            messagebox.showinfo("保存完了", "設定を保存しました！")
        except Exception as e:
            logger.error(f"設定保存エラー: {e}", exc_info=True)
            messagebox.showerror("保存エラー", f"設定の保存に失敗しました。\n\n詳細: {e}")

    def reload_settings(self) -> None:
        """設定を再読み込み"""
        self.on_reload()
        self._update_gs_status_sync()

    def open_config_file(self) -> None:
        """config.jsonをテキストエディタで開く"""
        config_path = self.config.config_path

        if open_file_or_folder(config_path, self._show_file_open_error):
            self.update_status("config.jsonを開きました")

    def _auto_detect_ghostscript(self) -> None:
        """Ghostscriptを自動検出（バックグラウンド実行でUIフリーズを防止）"""
        self.update_status("Ghostscriptを検索中...")
        self.gs_status_label.config(text="🔍 検索中...", fg="blue")

        def detect_task() -> None:
            from ghostscript_utils import GhostscriptManager

            gs_path = GhostscriptManager.find_ghostscript()
            verified = gs_path and GhostscriptManager.verify_ghostscript(gs_path)

            def update_ui() -> None:
                if verified:
                    self.gs_var.set(gs_path)
                    self._update_gs_status_sync()
                    self.update_status(f"Ghostscriptを検出: {gs_path}")
                    messagebox.showinfo("検出成功", f"Ghostscriptを検出しました。\n\n{gs_path}")
                else:
                    self._update_gs_status_sync()
                    instructions = GhostscriptManager.get_install_instructions()
                    messagebox.showwarning("未検出", instructions)

            try:
                self.tab.after(0, update_ui)
            except tk.TclError:
                pass

        thread = threading.Thread(target=detect_task, daemon=True)
        thread.start()

    def _check_gs_path(self, gs_path: str, verified: Optional[bool] = None) -> tuple:
        """
        Ghostscriptパスの状態を判定

        Args:
            gs_path: Ghostscriptの実行ファイルパス
            verified: 動作確認結果（Noneの場合はパス存在のみチェック）

        Returns:
            tuple: (表示テキスト, 色)
        """
        if not gs_path:
            return ("⚠️ 未設定（PDF圧縮機能は使用できません）", "orange")
        if gs_path.startswith('\\\\'):
            return ("⚠️ ネットワークパスは推奨されません", "orange")
        if not os.path.exists(gs_path):
            return ("❌ ファイルが存在しません", "red")
        if verified is None:
            return ("⏳ 動作確認中...", "gray")
        if verified:
            return ("✅ 正常に動作しています", "green")
        return ("❌ 動作確認に失敗しました", "red")

    def _update_gs_status(self) -> None:
        """Ghostscriptのステータスを更新（起動時用：subprocessを避けてUIフリーズを防止）"""
        gs_path = self.gs_var.get().strip()
        text, color = self._check_gs_path(gs_path)
        self.gs_status_label.config(text=text, fg=color)

        # パスが存在する場合は動作確認をバックグラウンドで実行
        if text == "⏳ 動作確認中...":
            self.tab.after(500, self._verify_gs_async)

    def _verify_gs_async(self) -> None:
        """Ghostscriptの動作確認をバックグラウンドで実行"""
        def verify_task() -> None:
            from ghostscript_utils import GhostscriptManager
            gs_path = self.gs_var.get().strip()
            verified = gs_path and GhostscriptManager.verify_ghostscript(gs_path)
            text, color = self._check_gs_path(gs_path, verified)

            thread_safe_call(self.tab, lambda: self.gs_status_label.config(text=text, fg=color))

        thread = threading.Thread(target=verify_task, daemon=True)
        thread.start()

    def _update_gs_status_sync(self) -> None:
        """Ghostscriptのステータスを同期的に更新（ユーザー操作後の即時反映用）"""
        from ghostscript_utils import GhostscriptManager
        gs_path = self.gs_var.get().strip()
        verified = GhostscriptManager.verify_ghostscript(gs_path) if os.path.exists(gs_path) and gs_path else None
        text, color = self._check_gs_path(gs_path, verified)
        self.gs_status_label.config(text=text, fg=color)

    def _test_ichitaro_conversion(self) -> None:
        """一太郎変換をテスト"""

        # jtdファイルを選択
        file_path = filedialog.askopenfilename(
            title="テスト用の一太郎ファイルを選択",
            filetypes=[("一太郎ファイル", "*.jtd"), ("すべて", "*.*")]
        )
        if not file_path:
            return

        self.ichitaro_status_label.config(text="🔄 テスト実行中...", fg="blue")
        self.tab.update()

        def run_test():
            try:
                from pdf_converter import PDFConverter
                import tempfile

                # 現在の設定を使用
                ichitaro_settings = self.config.get('ichitaro') or {}
                try:
                    ichitaro_settings['max_retries'] = int(self.max_retries_var.get())
                except ValueError:
                    pass
                try:
                    ichitaro_settings['save_wait_seconds'] = int(self.save_wait_var.get())
                except ValueError:
                    pass

                # セキュアな一時ファイル作成（TOCTOU攻撃対策）
                import uuid

                temp_dir = tempfile.gettempdir()
                converter = PDFConverter(temp_dir, ichitaro_settings)

                # UUID使用で衝突回避 + 安全なパス構築
                unique_id = uuid.uuid4().hex
                output_path = os.path.join(temp_dir, f"ichitaro_test_{unique_id}.pdf")

                try:
                    # 一時ファイルをプレースホルダとして作成（排他的作成）
                    fd = os.open(output_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
                    os.close(fd)  # ファイルディスクリプタを閉じて変換処理に渡す

                    result = converter._convert_ichitaro(file_path, output_path)

                    if result and os.path.exists(result):
                        thread_safe_call(self.tab, lambda: self.ichitaro_status_label.config(
                            text="✅ 変換成功！", fg="green"))
                        thread_safe_call(self.tab, lambda: messagebox.showinfo(
                            "テスト成功",
                            f"一太郎変換が成功しました。\n\n出力ファイル:\n{result}"
                        ))
                    else:
                        thread_safe_call(self.tab, lambda: self.ichitaro_status_label.config(
                            text="❌ 変換失敗", fg="red"))
                        thread_safe_call(self.tab, lambda: messagebox.showwarning(
                            "テスト失敗",
                            "一太郎変換に失敗しました。\n\n"
                            "リトライ回数の設定を調整してください。"
                        ))
                finally:
                    # 一時ファイルのクリーンアップ
                    try:
                        os.unlink(output_path)
                    except FileNotFoundError:
                        pass  # 既に削除済み
                    except Exception as cleanup_error:
                        logger.warning(f"一時ファイル削除失敗: {cleanup_error}")

            except Exception as test_error:
                error_msg = str(test_error)
                error_preview = error_msg[:50]
                thread_safe_call(self.tab, lambda: self.ichitaro_status_label.config(
                    text=f"❌ エラー: {error_preview}", fg="red"))
                thread_safe_call(self.tab, lambda: messagebox.showerror(
                    "テストエラー", f"テスト中にエラーが発生しました。\n\n{error_msg}"
                ))

        thread = threading.Thread(target=run_test, daemon=True)
        thread.start()

    def _open_log_file(self) -> None:
        """ログファイルを開く"""
        from datetime import datetime

        # ログディレクトリのパス
        appdata = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
        log_dir = os.path.join(appdata, 'PDFMergeSystem', 'logs')

        # 今日のログファイル
        log_file = os.path.join(log_dir, f"pdf_merge_{datetime.now():%Y%m%d}.log")

        if os.path.exists(log_file):
            # ログファイルをデフォルトのテキストエディタで開く
            if open_file_or_folder(log_file, self._show_file_open_error):
                self.update_status("ログファイルを開きました")
        else:
            # ログファイルが存在しない場合はログディレクトリを開く
            if os.path.exists(log_dir):
                if open_file_or_folder(log_dir, self._show_file_open_error):
                    self.update_status("ログディレクトリを開きました")
            else:
                messagebox.showwarning(
                    "ログファイルなし",
                    "ログファイルが見つかりません。\n\nまだ処理が実行されていない可能性があります。"
                )
