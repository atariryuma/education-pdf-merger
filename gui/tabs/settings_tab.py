"""
設定タブ

アプリケーション設定のUIを提供
"""
import logging
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Any, Callable, Optional, TYPE_CHECKING
from pathlib import Path

from gui.tabs.base_tab import BaseTab

if TYPE_CHECKING:
    from infrastructure.config_loader import ConfigLoader
from gui.event_names_editor import EventNamesEditor
from gui.styles import COLORS, FONTS
from gui.utils import create_hover_button, open_file_or_folder, thread_safe_call
from infrastructure.path_validator import PathValidator
from infrastructure.year_utils import calculate_year_short

logger = logging.getLogger(__name__)


class SettingsTab(BaseTab):
    """設定タブ"""

    def __init__(
        self,
        notebook: ttk.Notebook,
        config: "ConfigLoader",
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
        self.add_to_notebook("設定")

    def _on_year_changed(self, *args: Any) -> None:
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
        # ボタンを先に下部に固定（行事名展開時も常に見える）
        self._create_bottom_buttons()

        main_container = tk.Frame(self.tab)
        main_container.pack(fill="both", expand=True)

        # 共通のラベル幅とパディング
        LABEL_WIDTH = 16
        PAD_Y = 3

        # --- 年度情報 ---
        year_frame = tk.LabelFrame(main_container, text="年度情報", font=FONTS['default_bold'])
        year_frame.pack(fill="x", pady=(0, 4))

        tk.Label(year_frame, text="年度（西暦）:", width=LABEL_WIDTH, anchor="e").grid(row=0, column=0, sticky="e", padx=(10, 3), pady=PAD_Y)
        tk.Entry(year_frame, textvariable=self.year_var, width=15).grid(row=0, column=1, sticky="w", padx=3, pady=PAD_Y)
        tk.Label(year_frame, text="→", font=FONTS['default']).grid(row=0, column=2, sticky="w", padx=3, pady=PAD_Y)
        tk.Label(year_frame, textvariable=self.year_short_var, font=FONTS['default_bold'], fg=COLORS['primary']).grid(row=0, column=3, sticky="w", padx=3, pady=PAD_Y)
        tk.Label(year_frame, text="(自動計算)", font=FONTS['tiny'], fg="gray").grid(row=0, column=4, sticky="w", padx=3, pady=PAD_Y)

        # --- パス設定 ---
        path_frame = tk.LabelFrame(main_container, text="パス設定", font=FONTS['default_bold'])
        path_frame.pack(fill="x", pady=4)

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
        tool_frame = tk.LabelFrame(main_container, text="ツール設定", font=FONTS['default_bold'])
        tool_frame.pack(fill="x", pady=4)

        tk.Label(tool_frame, text="Ghostscript:", width=LABEL_WIDTH, anchor="e").grid(row=0, column=0, sticky="e", padx=(10, 3), pady=PAD_Y)
        tk.Entry(tool_frame, textvariable=self.gs_var).grid(row=0, column=1, sticky="ew", padx=3, pady=PAD_Y)

        gs_btn_frame = tk.Frame(tool_frame)
        gs_btn_frame.grid(row=0, column=2, padx=(3, 10), pady=PAD_Y)
        tk.Button(gs_btn_frame, text="📄", command=self._browse_gs_file, width=3).pack(side="left", padx=1)
        tk.Button(gs_btn_frame, text="🔍 自動検出", command=self._auto_detect_ghostscript, font=FONTS['tiny']).pack(side="left", padx=1)
        tk.Button(gs_btn_frame, text="⬇ インストール", command=self._install_ghostscript, font=FONTS['tiny']).pack(side="left", padx=1)

        # Ghostscriptステータス表示
        self.gs_status_label = tk.Label(tool_frame, text="", fg="gray", font=FONTS['tiny'])
        self.gs_status_label.grid(row=1, column=1, columnspan=2, sticky="w", padx=3, pady=(0, 3))
        self._update_gs_status()

        tool_frame.columnconfigure(1, weight=1)

        # --- 一太郎設定 ---
        ichitaro_frame = tk.LabelFrame(main_container, text="一太郎変換設定", font=FONTS['default_bold'])
        ichitaro_frame.pack(fill="x", pady=(0, 4))

        max_retries = self.config.get('ichitaro', 'max_retries')
        self.max_retries_var = tk.StringVar(value=str(max_retries if max_retries is not None else 3))
        save_wait = self.config.get('ichitaro', 'save_wait_seconds')
        self.save_wait_var = tk.StringVar(value=str(save_wait if save_wait is not None else 20))

        settings_row = tk.Frame(ichitaro_frame)
        settings_row.pack(fill="x", padx=10, pady=4)
        tk.Label(settings_row, text="リトライ:").pack(side="left")
        tk.Entry(settings_row, textvariable=self.max_retries_var, width=3).pack(side="left", padx=(3, 0))
        tk.Label(settings_row, text="回").pack(side="left", padx=(2, 15))
        tk.Label(settings_row, text="保存待機:").pack(side="left")
        tk.Entry(settings_row, textvariable=self.save_wait_var, width=3).pack(side="left", padx=(3, 0))
        tk.Label(settings_row, text="秒").pack(side="left", padx=(2, 15))
        tk.Button(settings_row, text="テスト", command=self._test_ichitaro_conversion, font=FONTS['tiny']).pack(side="left", padx=5)
        self.ichitaro_status_label = tk.Label(settings_row, text="", fg="gray", font=FONTS['tiny'])
        self.ichitaro_status_label.pack(side="left", padx=5)

        # --- 行事名設定 ---
        event_frame = tk.LabelFrame(main_container, text="行事名設定（Excel転記用）", font=FONTS['default_bold'])
        event_frame.pack(fill="both", expand=True, pady=4)

        self.event_names_editor = EventNamesEditor(
            event_frame, self.config, status_callback=self.update_status
        )

    def _create_bottom_buttons(self) -> None:
        """ボタン行を作成（タブ下部に固定）"""
        button_frame = tk.Frame(self.tab)
        button_frame.pack(side="bottom", fill="x", pady=10)

        # 中央寄せ用の内部フレーム
        inner = tk.Frame(button_frame)
        inner.pack()

        create_hover_button(
            inner, text="💾 保存 (Ctrl+S)", command=self.save_settings,
            color="primary", font=FONTS['small_bold'], width=14, height=1
        ).pack(side="left", padx=5)

        tk.Button(
            inner, text="🔄 再読み込み (Ctrl+R)", command=self.reload_settings,
            font=FONTS['small'], width=18, height=1, cursor="hand2"
        ).pack(side="left", padx=5)

        tk.Button(
            inner, text="📝 config.json編集", command=self.open_config_file,
            font=FONTS['small'], width=16, height=1, cursor="hand2"
        ).pack(side="left", padx=5)

    def reload_event_names(self) -> None:
        """すべてのカテゴリの行事名をリロード（後方互換性のための委譲メソッド）"""
        self.event_names_editor.reload_event_names()

    def _browse_folder(self, var: tk.StringVar) -> None:
        """フォルダを参照（PathValidatorベース）"""
        try:
            current_path_str = var.get().strip()
            initial_dir = PathValidator.get_safe_initial_dir(current_path_str, Path.home())

            directory = filedialog.askdirectory(title="フォルダを選択", initialdir=str(initial_dir))
            if directory:
                validated_path = self.validate_path(directory, "directory")
                if validated_path:
                    var.set(str(validated_path))
                    self.update_status(f"フォルダを選択: {validated_path.name}")
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
                validated_path = self.validate_path(file_path, "file", error_title="パス検証エラー")
                if not validated_path:
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
        except Exception as e:
            logger.error(f"設定保存エラー: {e}", exc_info=True)
            messagebox.showerror("保存エラー", f"設定の保存に失敗しました。\n\n詳細: {e}")

    def reload_settings(self) -> None:
        """設定を再読み込み"""
        self.on_reload()
        self._update_gs_status()  # 非同期で検証（UIフリーズ防止）

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
            from infrastructure.ghostscript import GhostscriptDetector

            gs_path = GhostscriptDetector.detect()
            verified = gs_path and GhostscriptDetector.verify(gs_path)

            def update_ui() -> None:
                if verified:
                    self.gs_var.set(gs_path)
                    self._update_gs_status_sync()
                    self.update_status(f"Ghostscriptを検出: {gs_path}")
                else:
                    self._update_gs_status_sync()
                    instructions = GhostscriptDetector.get_install_instructions()
                    messagebox.showwarning("未検出", instructions)

            try:
                self.tab.after(0, update_ui)
            except tk.TclError:
                pass

        thread = threading.Thread(target=detect_task, daemon=True)
        thread.start()

    def _install_ghostscript(self) -> None:
        """Ghostscriptをダウンロードしてインストール"""
        if self.gs_var.get().strip():
            result = messagebox.askyesno(
                "確認", "Ghostscriptは既に設定されています。\n再インストールしますか？"
            )
            if not result:
                return

        self.gs_status_label.config(text="⬇ ダウンロード中...", fg="blue")
        self.update_status("Ghostscriptをダウンロード中...")

        def install_task() -> None:
            from infrastructure.ghostscript import GhostscriptInstaller

            success, message, gs_path = GhostscriptInstaller.download_and_install(
                progress_callback=lambda msg: self.tab.after(
                    0, lambda m=msg: self.gs_status_label.config(text=m, fg="blue")
                )
            )

            def update_ui() -> None:
                if success and gs_path:
                    self.gs_var.set(gs_path)
                    self._update_gs_status_sync()
                    self.update_status(f"Ghostscriptインストール完了: {gs_path}")
                    messagebox.showinfo("完了", message)
                else:
                    self._update_gs_status_sync()
                    self.update_status("Ghostscriptインストール失敗")
                    messagebox.showerror("エラー", message)

            try:
                self.tab.after(0, update_ui)
            except tk.TclError:
                pass

        thread = threading.Thread(target=install_task, daemon=True)
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

        # パスが存在し未検証の場合は動作確認をバックグラウンドで実行
        if gs_path and not gs_path.startswith('\\\\') and os.path.exists(gs_path):
            self.tab.after(500, self._verify_gs_async)

    def _verify_gs_async(self) -> None:
        """Ghostscriptの動作確認をバックグラウンドで実行"""
        def verify_task() -> None:
            from infrastructure.ghostscript import GhostscriptDetector
            gs_path = self.gs_var.get().strip()
            verified = gs_path and GhostscriptDetector.verify(gs_path)
            text, color = self._check_gs_path(gs_path, verified)

            thread_safe_call(self.tab, lambda: self.gs_status_label.config(text=text, fg=color))

        thread = threading.Thread(target=verify_task, daemon=True)
        thread.start()

    def _update_gs_status_sync(self) -> None:
        """Ghostscriptのステータスを同期的に更新（ユーザー操作後の即時反映用）"""
        from infrastructure.ghostscript import GhostscriptDetector
        gs_path = self.gs_var.get().strip()
        verified = GhostscriptDetector.verify(gs_path) if os.path.exists(gs_path) and gs_path else None
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
        self.tab.update_idletasks()

        def run_test():
            try:
                from core.pdf_converter import PDFConverter
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

                    result = converter.convert(file_path, output_path)

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

