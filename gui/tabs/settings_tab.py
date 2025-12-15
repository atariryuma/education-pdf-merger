"""
設定タブ

アプリケーション設定のUIを提供
"""
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Any, Callable

from gui.tabs.base_tab import BaseTab
from gui.utils import create_hover_button


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
        network_var: tk.StringVar,
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
        self.network_var = network_var
        self.temp_var = temp_var
        self.gs_var = gs_var
        self.excel_ref_var = excel_ref_var
        self.excel_target_var = excel_target_var
        self.on_reload = on_reload
        self._create_ui()
        self.add_to_notebook("⚙️ 設定")

    def _create_ui(self) -> None:
        """UIを構築"""
        # メインコンテナ（中央配置用）
        main_container = tk.Frame(self.tab)
        main_container.pack(fill="both", expand=True, padx=15, pady=10)

        # 共通のラベル幅とパディング
        LABEL_WIDTH = 16
        PAD_Y = 5

        # --- 年度情報 ---
        year_frame = tk.LabelFrame(main_container, text="📅 年度情報", font=("メイリオ", 10, "bold"))
        year_frame.pack(fill="x", pady=(0, 8))

        tk.Label(year_frame, text="年度（フル）:", width=LABEL_WIDTH, anchor="e").grid(row=0, column=0, sticky="e", padx=(10, 3), pady=PAD_Y)
        tk.Entry(year_frame, textvariable=self.year_var, width=25).grid(row=0, column=1, sticky="w", padx=3, pady=PAD_Y)
        tk.Label(year_frame, text="略称:", anchor="e").grid(row=0, column=2, sticky="e", padx=(10, 3), pady=PAD_Y)
        tk.Entry(year_frame, textvariable=self.year_short_var, width=8).grid(row=0, column=3, sticky="w", padx=(3, 10), pady=PAD_Y)

        # --- パス設定 ---
        path_frame = tk.LabelFrame(main_container, text="📂 パス設定", font=("メイリオ", 10, "bold"))
        path_frame.pack(fill="x", pady=8)

        tk.Label(path_frame, text="Google Drive:", width=LABEL_WIDTH, anchor="e").grid(row=0, column=0, sticky="e", padx=(10, 3), pady=PAD_Y)
        tk.Entry(path_frame, textvariable=self.gdrive_var).grid(row=0, column=1, sticky="ew", padx=3, pady=PAD_Y)
        tk.Button(path_frame, text="📁", command=lambda: self._browse_folder(self.gdrive_var), width=3).grid(row=0, column=2, padx=(3, 10), pady=PAD_Y)

        tk.Label(path_frame, text="ネットワーク:", width=LABEL_WIDTH, anchor="e").grid(row=1, column=0, sticky="e", padx=(10, 3), pady=PAD_Y)
        tk.Entry(path_frame, textvariable=self.network_var).grid(row=1, column=1, sticky="ew", padx=3, pady=PAD_Y)
        tk.Button(path_frame, text="📁", command=lambda: self._browse_folder(self.network_var), width=3).grid(row=1, column=2, padx=(3, 10), pady=PAD_Y)

        tk.Label(path_frame, text="一時フォルダ:", width=LABEL_WIDTH, anchor="e").grid(row=2, column=0, sticky="e", padx=(10, 3), pady=PAD_Y)
        tk.Entry(path_frame, textvariable=self.temp_var).grid(row=2, column=1, sticky="ew", padx=3, pady=PAD_Y)

        temp_btn_frame = tk.Frame(path_frame)
        temp_btn_frame.grid(row=2, column=2, padx=(3, 10), pady=PAD_Y)
        tk.Button(temp_btn_frame, text="📁", command=lambda: self._browse_folder(self.temp_var), width=3).pack(side="left", padx=1)
        tk.Button(temp_btn_frame, text="📂 開く", command=self._open_temp_folder, font=("メイリオ", 8)).pack(side="left", padx=1)

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
        self.down_arrow_var = tk.StringVar(value=str(self.config.get('ichitaro', 'down_arrow_count') or 5))

        # 1行目: リトライ回数と保存待機時間
        settings_row1 = tk.Frame(ichitaro_frame)
        settings_row1.pack(fill="x", padx=10, pady=PAD_Y)
        tk.Label(settings_row1, text="リトライ:").pack(side="left")
        tk.Entry(settings_row1, textvariable=self.max_retries_var, width=3).pack(side="left", padx=(3, 0))
        tk.Label(settings_row1, text="回").pack(side="left", padx=(2, 15))
        tk.Label(settings_row1, text="保存待機:").pack(side="left")
        tk.Entry(settings_row1, textvariable=self.save_wait_var, width=3).pack(side="left", padx=(3, 0))
        tk.Label(settings_row1, text="秒").pack(side="left", padx=(2, 15))
        tk.Button(settings_row1, text="🧪 テスト", command=self._test_ichitaro_conversion, font=("メイリオ", 8)).pack(side="left", padx=5)

        # 2行目: 下矢印回数（プリンタ選択）
        settings_row2 = tk.Frame(ichitaro_frame)
        settings_row2.pack(fill="x", padx=10, pady=PAD_Y)
        tk.Label(settings_row2, text="↓回数:").pack(side="left")
        tk.Entry(settings_row2, textvariable=self.down_arrow_var, width=3).pack(side="left", padx=(3, 0))
        tk.Label(settings_row2, text="回").pack(side="left", padx=(2, 5))
        tk.Label(settings_row2, text="（Microsoft Print to PDFまでの下矢印キー押下回数）", fg="#666", font=("メイリオ", 8)).pack(side="left")

        # 説明ラベル
        help_label = tk.Label(
            ichitaro_frame,
            text="💡 ヒント: プリンタの並び順が変わった場合は「↓回数」を調整してください。",
            fg="#0066cc",
            font=("メイリオ", 8)
        )
        help_label.pack(anchor="w", padx=10, pady=(0, 3))

        # ステータス表示
        self.ichitaro_status_label = tk.Label(ichitaro_frame, text="処理手順: Ctrl+P → ↓キー×N回 → Enter → ファイル名 → Enter", fg="#666", font=("メイリオ", 8))
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
        tk.Button(excel_frame, text="📄", command=lambda: self._browse_excel_file(self.excel_ref_var), width=3).grid(row=0, column=2, padx=(3, 10), pady=PAD_Y)

        tk.Label(excel_frame, text="対象:", width=LABEL_WIDTH, anchor="e").grid(row=1, column=0, sticky="e", padx=(10, 3), pady=PAD_Y)
        tk.Entry(excel_frame, textvariable=self.excel_target_var).grid(row=1, column=1, sticky="ew", padx=3, pady=PAD_Y)
        tk.Button(excel_frame, text="📄", command=lambda: self._browse_excel_file(self.excel_target_var), width=3).grid(row=1, column=2, padx=(3, 10), pady=PAD_Y)

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
        """フォルダを参照"""
        try:
            current_path = var.get().strip()
            if current_path and os.path.exists(current_path) and os.path.isdir(current_path):
                initial_dir = current_path
            elif current_path and os.path.dirname(current_path) and os.path.exists(os.path.dirname(current_path)):
                initial_dir = os.path.dirname(current_path)
            else:
                initial_dir = os.path.expanduser("~")

            directory = filedialog.askdirectory(title="フォルダを選択", initialdir=initial_dir)
            if directory:
                var.set(directory)
                self.update_status(f"フォルダを選択: {os.path.basename(directory)}")
        except Exception as e:
            messagebox.showerror("参照エラー", f"フォルダの参照中にエラーが発生しました。\n\n詳細: {e}")

    def _browse_gs_file(self) -> None:
        """Ghostscript実行ファイルを参照"""
        try:
            current_path = self.gs_var.get().strip()
            if current_path and os.path.exists(current_path) and os.path.isfile(current_path):
                initial_dir = os.path.dirname(current_path)
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
                self.update_status(f"Ghostscript: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("参照エラー", f"ファイルの参照中にエラーが発生しました。\n\n詳細: {e}")

    def _browse_excel_file(self, var: tk.StringVar) -> None:
        """Excelファイルを参照"""
        try:
            base_path = self.config.get('base_paths', 'google_drive')
            year = self.config.year
            education_base = self.config.get('directories', 'education_plan_base')
            initial_dir = os.path.join(base_path, year, education_base)

            if not os.path.exists(initial_dir):
                initial_dir = os.path.expanduser("~")

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

    def _open_temp_folder(self) -> None:
        """一時フォルダをエクスプローラーで開く"""
        temp_path = self.temp_var.get().strip()

        # パスが空の場合はデフォルトパスを使用
        if not temp_path:
            appdata = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
            temp_path = os.path.join(appdata, 'PDFMergeSystem', 'temp')

        # フォルダが存在しない場合は作成
        if not os.path.exists(temp_path):
            try:
                os.makedirs(temp_path)
                self.update_status(f"一時フォルダを作成しました: {temp_path}")
            except Exception as e:
                messagebox.showerror("エラー", f"一時フォルダの作成に失敗しました。\n\n{e}")
                return

        # エクスプローラーで開く
        try:
            os.startfile(temp_path)
            self.update_status(f"一時フォルダを開きました")
        except Exception as e:
            messagebox.showerror("エラー", f"フォルダを開けませんでした。\n\n{e}")

    def save_settings(self) -> None:
        """設定を保存"""
        year = self.year_var.get().strip()
        year_short = self.year_short_var.get().strip()

        if not year or not year_short:
            messagebox.showerror("入力エラー", "年度情報は必須です。")
            return

        self.config.update_year(year, year_short)
        self.config.set('base_paths', 'google_drive', value=self.gdrive_var.get())
        self.config.set('base_paths', 'network', value=self.network_var.get())
        self.config.set('base_paths', 'local_temp', value=self.temp_var.get())
        self.config.set('ghostscript', 'executable', value=self.gs_var.get())
        self.config.set('files', 'excel_reference', value=self.excel_ref_var.get())
        self.config.set('files', 'excel_target', value=self.excel_target_var.get())

        # 一太郎設定の保存
        try:
            self.config.set('ichitaro', 'max_retries', value=int(self.max_retries_var.get()))
        except ValueError:
            pass
        try:
            self.config.set('ichitaro', 'save_wait_seconds', value=int(self.save_wait_var.get()))
        except ValueError:
            pass
        try:
            self.config.set('ichitaro', 'down_arrow_count', value=int(self.down_arrow_var.get()))
        except ValueError:
            pass

        if self.config.save_config():
            self.update_status("設定を保存しました")
            messagebox.showinfo("✅ 保存完了", "設定を保存しました！")
        else:
            messagebox.showerror("❌ 保存エラー", "設定の保存に失敗しました。")

    def reload_settings(self) -> None:
        """設定を再読み込み"""
        self.on_reload()

    def open_config_file(self) -> None:
        """config.jsonをテキストエディタで開く"""
        config_path = self.config.config_path
        if os.path.exists(config_path):
            os.startfile(config_path)
            self.update_status(f"config.jsonを開きました")
        else:
            messagebox.showerror("❌ ファイルエラー", f"config.jsonが見つかりません。\n\nパス: {config_path}")

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
            messagebox.showinfo("✅ 検出成功", f"Ghostscriptを検出しました。\n\n{gs_path}")
        else:
            self._update_gs_status()
            instructions = GhostscriptManager.get_install_instructions()
            messagebox.showwarning("⚠️ 未検出", instructions)

    def _update_gs_status(self) -> None:
        """Ghostscriptのステータスを更新"""
        from ghostscript_utils import GhostscriptManager

        gs_path = self.gs_var.get().strip()

        if not gs_path:
            self.gs_status_label.config(text="⚠️ 未設定（PDF圧縮機能は使用できません）", fg="orange")
        elif not os.path.exists(gs_path):
            self.gs_status_label.config(text="❌ ファイルが存在しません", fg="red")
        elif GhostscriptManager.verify_ghostscript(gs_path):
            self.gs_status_label.config(text="✅ 正常に動作しています", fg="green")
        else:
            self.gs_status_label.config(text="❌ 動作確認に失敗しました", fg="red")

    def _test_ichitaro_conversion(self) -> None:
        """一太郎変換をテスト"""
        from tkinter import filedialog
        import threading

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
                try:
                    ichitaro_settings['down_arrow_count'] = int(self.down_arrow_var.get())
                except ValueError:
                    pass

                # 一時ディレクトリを使用
                temp_dir = tempfile.gettempdir()
                converter = PDFConverter(temp_dir, ichitaro_settings)

                output_path = os.path.join(temp_dir, "ichitaro_test_output.pdf")
                if os.path.exists(output_path):
                    os.remove(output_path)

                result = converter._convert_ichitaro(file_path, output_path)

                if result and os.path.exists(result):
                    self.tab.after(0, lambda: self.ichitaro_status_label.config(
                        text=f"✅ 変換成功！", fg="green"))
                    self.tab.after(0, lambda: messagebox.showinfo(
                        "✅ テスト成功",
                        f"一太郎変換が成功しました。\n\n出力ファイル:\n{result}"
                    ))
                else:
                    self.tab.after(0, lambda: self.ichitaro_status_label.config(
                        text="❌ 変換失敗", fg="red"))
                    self.tab.after(0, lambda: messagebox.showwarning(
                        "⚠️ テスト失敗",
                        "一太郎変換に失敗しました。\n\n"
                        "「↓回数」の設定を調整してください。"
                    ))
            except Exception as e:
                self.tab.after(0, lambda: self.ichitaro_status_label.config(
                    text=f"❌ エラー: {str(e)[:50]}", fg="red"))
                error_msg = str(e)
                self.tab.after(0, lambda: messagebox.showerror(
                    "❌ テストエラー", f"テスト中にエラーが発生しました。\n\n{error_msg}"
                ))

        thread = threading.Thread(target=run_test, daemon=True)
        thread.start()

    def _open_log_file(self) -> None:
        """ログファイルを開く"""
        import os
        from datetime import datetime

        # ログディレクトリのパス
        appdata = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
        log_dir = os.path.join(appdata, 'PDFMergeSystem', 'logs')

        # 今日のログファイル
        log_file = os.path.join(log_dir, f"pdf_merge_{datetime.now():%Y%m%d}.log")

        if os.path.exists(log_file):
            # ログファイルをデフォルトのテキストエディタで開く
            os.startfile(log_file)
        else:
            # ログファイルが存在しない場合はログディレクトリを開く
            if os.path.exists(log_dir):
                os.startfile(log_dir)
            else:
                messagebox.showwarning(
                    "⚠️ ログファイルなし",
                    f"ログファイルが見つかりません。\n\nまだ処理が実行されていない可能性があります。"
                )
