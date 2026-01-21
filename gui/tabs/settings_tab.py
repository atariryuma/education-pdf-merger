"""
設定タブ

アプリケーション設定のUIを提供
"""
import logging
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Any, Callable
from pathlib import Path

from gui.tabs.base_tab import BaseTab
from gui.utils import create_hover_button, open_file_or_folder
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
        excel_ref_var: tk.StringVar,
        excel_target_var: tk.StringVar,
        on_reload: Callable[[], None]
    ) -> None:
        super().__init__(notebook, config, status_bar)
        self.year_var = year_var
        self.year_short_var = year_short_var
        self.gdrive_var = gdrive_var
        self.temp_var = temp_var
        self.gs_var = gs_var
        self.excel_ref_var = excel_ref_var
        self.excel_target_var = excel_target_var
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

        # --- Excelファイル設定 ---
        excel_frame = tk.LabelFrame(main_container, text="📊 Excelファイル設定", font=("メイリオ", 10, "bold"))
        excel_frame.pack(fill="x", pady=8)

        tk.Label(excel_frame, text="参照元:", width=LABEL_WIDTH, anchor="e").grid(row=0, column=0, sticky="e", padx=(10, 3), pady=PAD_Y)
        tk.Entry(excel_frame, textvariable=self.excel_ref_var).grid(row=0, column=1, sticky="ew", padx=3, pady=PAD_Y)

        excel_ref_btn_frame = tk.Frame(excel_frame)
        excel_ref_btn_frame.grid(row=0, column=2, padx=(3, 10), pady=PAD_Y)
        tk.Button(excel_ref_btn_frame, text="📄", command=lambda: self._browse_excel_file(self.excel_ref_var), width=3).pack(side="left", padx=1)
        tk.Button(excel_ref_btn_frame, text="📂", command=lambda: self._open_excel_file_from_settings(self.excel_ref_var), width=3).pack(side="left", padx=1)

        tk.Label(excel_frame, text="対象:", width=LABEL_WIDTH, anchor="e").grid(row=1, column=0, sticky="e", padx=(10, 3), pady=PAD_Y)
        tk.Entry(excel_frame, textvariable=self.excel_target_var).grid(row=1, column=1, sticky="ew", padx=3, pady=PAD_Y)

        excel_target_btn_frame = tk.Frame(excel_frame)
        excel_target_btn_frame.grid(row=1, column=2, padx=(3, 10), pady=PAD_Y)
        tk.Button(excel_target_btn_frame, text="📄", command=lambda: self._browse_excel_file(self.excel_target_var), width=3).pack(side="left", padx=1)
        tk.Button(excel_target_btn_frame, text="📂", command=lambda: self._open_excel_file_from_settings(self.excel_target_var), width=3).pack(side="left", padx=1)

        excel_frame.columnconfigure(1, weight=1)

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
                self.gs_var.set(file_path)
                self._update_gs_status()
                self.update_status(f"Ghostscript: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("参照エラー", f"ファイルの参照中にエラーが発生しました。\n\n詳細: {e}")

    def _browse_excel_file(self, var: tk.StringVar) -> None:
        """Excelファイルを参照"""
        try:
            base_path = self.config.get('base_paths', 'google_drive')
            year = self.config.year
            year_short = self.config.year_short
            education_base = self.config.get('directories', 'education_plan_base')

            # {year_short}プレースホルダーを実際の値に置き換える
            education_base = education_base.replace('{year_short}', year_short)

            initial_dir_candidate = os.path.join(base_path, year, education_base)

            # フリーズ防止: PathValidator.get_safe_initial_dirを使用
            safe_initial_dir = PathValidator.get_safe_initial_dir(
                initial_dir_candidate,
                fallback=Path.home()
            )
            initial_dir = str(safe_initial_dir)

            file_path = filedialog.askopenfilename(
                title="Excelファイルを選択",
                initialdir=initial_dir,
                filetypes=[("Excel", "*.xlsx;*.xls"), ("すべて", "*.*")]
            )
            if file_path:
                var.set(os.path.basename(file_path))
                self.update_status(f"Excelファイル: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("参照エラー", f"Excelファイルの参照中にエラーが発生しました。\n\n詳細: {e}")

    def _open_folder(self, var: tk.StringVar) -> None:
        """フォルダをエクスプローラーで開く"""
        folder_path_str = var.get().strip()

        if not folder_path_str:
            messagebox.showwarning("警告", "フォルダパスが設定されていません。")
            return

        def on_error(error_msg: str):
            messagebox.showerror("エラー", error_msg)

        if open_file_or_folder(folder_path_str, on_error):
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
        def on_error(error_msg: str):
            messagebox.showerror("エラー", error_msg)

        if open_file_or_folder(str(temp_path), on_error):
            self.update_status("一時フォルダを開きました")

    def _open_excel_file_from_settings(self, var: tk.StringVar) -> None:
        """Excelファイルを開く（設定タブから）"""
        filename = var.get().strip()

        if not filename:
            messagebox.showwarning("警告", "ファイル名が設定されていません。")
            return

        base_path = self.config.get('base_paths', 'google_drive')
        year = self.config.year
        year_short = self.config.year_short
        education_base = self.config.get('directories', 'education_plan_base')

        # {year_short}プレースホルダーを実際の値に置き換える
        education_base = education_base.replace('{year_short}', year_short)

        file_path = os.path.join(base_path, year, education_base, filename)

        def on_error(error_msg: str):
            messagebox.showerror("エラー", error_msg)

        if open_file_or_folder(file_path, on_error):
            self.update_status(f"Excelファイルを開きました: {filename}")

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
        self.config.set('files', 'excel_reference', value=self.excel_ref_var.get())
        self.config.set('files', 'excel_target', value=self.excel_target_var.get())

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
        self._update_gs_status()

    def open_config_file(self) -> None:
        """config.jsonをテキストエディタで開く"""
        config_path = self.config.config_path

        def on_error(error_msg: str):
            messagebox.showerror("エラー", error_msg)

        if open_file_or_folder(config_path, on_error):
            self.update_status("config.jsonを開きました")

    def _auto_detect_ghostscript(self) -> None:
        """Ghostscriptを自動検出"""
        from ghostscript_utils import GhostscriptManager

        self.update_status("Ghostscriptを検索中...")
        self.gs_status_label.config(text="🔍 検索中...", fg="blue")
        self.tab.update()

        gs_path = GhostscriptManager.find_ghostscript()

        if gs_path and GhostscriptManager.verify_ghostscript(gs_path):
            self.gs_var.set(gs_path)
            self._update_gs_status()
            self.update_status(f"Ghostscriptを検出: {gs_path}")
            messagebox.showinfo("検出成功", f"Ghostscriptを検出しました。\n\n{gs_path}")
        else:
            self._update_gs_status()
            instructions = GhostscriptManager.get_install_instructions()
            messagebox.showwarning("未検出", instructions)

    def _update_gs_status(self) -> None:
        """Ghostscriptのステータスを更新（フリーズ防止版）"""
        from ghostscript_utils import GhostscriptManager

        gs_path = self.gs_var.get().strip()

        if not gs_path:
            self.gs_status_label.config(text="⚠️ 未設定（PDF圧縮機能は使用できません）", fg="orange")
        else:
            # ローカルパスのみチェック（ネットワークパスは警告）
            if gs_path.startswith('\\\\'):
                self.gs_status_label.config(text="⚠️ ネットワークパスは推奨されません", fg="orange")
            elif not os.path.exists(gs_path):
                self.gs_status_label.config(text="❌ ファイルが存在しません", fg="red")
            elif GhostscriptManager.verify_ghostscript(gs_path):
                self.gs_status_label.config(text="✅ 正常に動作しています", fg="green")
            else:
                self.gs_status_label.config(text="❌ 動作確認に失敗しました", fg="red")

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
                        self.tab.after(0, lambda: self.ichitaro_status_label.config(
                            text="✅ 変換成功！", fg="green"))
                        self.tab.after(0, lambda: messagebox.showinfo(
                            "テスト成功",
                            f"一太郎変換が成功しました。\n\n出力ファイル:\n{result}"
                        ))
                    else:
                        self.tab.after(0, lambda: self.ichitaro_status_label.config(
                            text="❌ 変換失敗", fg="red"))
                        self.tab.after(0, lambda: messagebox.showwarning(
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
                self.tab.after(0, lambda: self.ichitaro_status_label.config(
                    text=f"❌ エラー: {error_preview}", fg="red"))
                self.tab.after(0, lambda: messagebox.showerror(
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

        def on_error(error_msg: str):
            messagebox.showerror("エラー", error_msg)

        if os.path.exists(log_file):
            # ログファイルをデフォルトのテキストエディタで開く
            if open_file_or_folder(log_file, on_error):
                self.update_status("ログファイルを開きました")
        else:
            # ログファイルが存在しない場合はログディレクトリを開く
            if os.path.exists(log_dir):
                if open_file_or_folder(log_dir, on_error):
                    self.update_status("ログディレクトリを開きました")
            else:
                messagebox.showwarning(
                    "ログファイルなし",
                    "ログファイルが見つかりません。\n\nまだ処理が実行されていない可能性があります。"
                )
