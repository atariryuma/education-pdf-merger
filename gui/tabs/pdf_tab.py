"""
PDF統合タブ

PDF統合機能のUIを提供
"""
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
from typing import Any, Optional

from gui.tabs.base_tab import BaseTab
from gui.utils import set_button_state, create_hover_button, thread_safe_call


class PDFTab(BaseTab):
    """PDF統合タブ"""

    def __init__(
        self,
        notebook: ttk.Notebook,
        config: Any,
        status_bar: tk.Label,
        input_dir_var: tk.StringVar,
        output_file_var: tk.StringVar,
        plan_type_var: tk.StringVar
    ) -> None:
        super().__init__(notebook, config, status_bar)
        self.input_dir_var = input_dir_var
        self.output_file_var = output_file_var
        self.plan_type_var = plan_type_var
        # スレッドセーフなキャンセルフラグ（threading.Eventを使用）
        self._cancel_event = threading.Event()
        self._create_ui()
        self.add_to_notebook("📄 PDF統合")

    def _create_ui(self) -> None:
        """UIを構築"""
        # 入力フォームのフレーム
        form_frame = tk.Frame(self.tab)
        form_frame.pack(fill="x", padx=20, pady=15)

        LABEL_WIDTH = 18

        # 入力ディレクトリ選択
        tk.Label(form_frame, text="入力ディレクトリ:", width=LABEL_WIDTH, anchor="e").grid(row=0, column=0, sticky="e", padx=(15, 5), pady=6)
        tk.Entry(form_frame, textvariable=self.input_dir_var).grid(row=0, column=1, padx=5, pady=6, sticky="ew")
        tk.Button(form_frame, text="📁 参照", command=self._select_input_dir, width=8).grid(row=0, column=2, padx=(5, 15), pady=6)

        # 出力ファイル選択
        tk.Label(form_frame, text="出力ファイル:", width=LABEL_WIDTH, anchor="e").grid(row=1, column=0, sticky="e", padx=(15, 5), pady=6)
        tk.Entry(form_frame, textvariable=self.output_file_var).grid(row=1, column=1, padx=5, pady=6, sticky="ew")
        tk.Button(form_frame, text="💾 参照", command=self._select_output_file, width=8).grid(row=1, column=2, padx=(5, 15), pady=6)

        # 計画種別選択
        tk.Label(form_frame, text="計画種別:", width=LABEL_WIDTH, anchor="e").grid(row=2, column=0, sticky="e", padx=(15, 5), pady=6)
        plan_frame = tk.Frame(form_frame)
        plan_frame.grid(row=2, column=1, sticky="w", padx=5, pady=6)
        tk.Radiobutton(
            plan_frame, text="📚 教育計画", variable=self.plan_type_var,
            value="education", font=("メイリオ", 10)
        ).pack(side="left", padx=(0, 15))
        tk.Radiobutton(
            plan_frame, text="📅 行事計画", variable=self.plan_type_var,
            value="event", font=("メイリオ", 10)
        ).pack(side="left", padx=15)

        form_frame.columnconfigure(1, weight=1)

        # 実行ボタン
        button_frame = tk.Frame(self.tab)
        button_frame.pack(pady=15)

        self.run_button = create_hover_button(
            button_frame,
            text="▶ PDF統合を実行",
            command=self._run_pdf_merge,
            color="primary",
            font=("メイリオ", 11, "bold"),
            width=28,
            height=2
        )
        self.run_button.pack(side="left", padx=5)

        self.cancel_button = tk.Button(
            button_frame,
            text="✕ キャンセル",
            command=self._cancel_operation,
            font=("メイリオ", 10),
            bg="#f44336",
            fg="white",
            width=12,
            height=2,
            state="disabled"
        )
        self.cancel_button.pack(side="left", padx=5)

        # ステータスラベル
        self.status_label = tk.Label(self.tab, text="", font=("メイリオ", 9), fg="gray")
        self.status_label.pack()

        # プログレスバー
        self.progress = ttk.Progressbar(self.tab, mode='indeterminate')
        self.progress.pack(fill="x", padx=20, pady=5)

        # ログ表示
        self.create_log_frame(height=10)
        # GUIログハンドラを設定（各モジュールのログをGUIに表示）
        self.setup_gui_logging()
        self.log("準備完了。入力ディレクトリと出力ファイルを選択して実行してください。", "info")

    def _select_input_dir(self) -> None:
        """入力ディレクトリを選択"""
        try:
            current_path = self.input_dir_var.get().strip()
            if current_path and os.path.exists(current_path) and os.path.isdir(current_path):
                initial_dir = current_path
            elif current_path and os.path.dirname(current_path) and os.path.exists(os.path.dirname(current_path)):
                initial_dir = os.path.dirname(current_path)
            else:
                default_input = self.config.get_education_plan_path()
                initial_dir = default_input if os.path.exists(default_input) else os.path.expanduser("~")

            directory = filedialog.askdirectory(title="入力ディレクトリを選択", initialdir=initial_dir)
            if directory:
                self.input_dir_var.set(directory)
                self.update_status(f"入力ディレクトリを選択: {os.path.basename(directory)}")
        except Exception as e:
            messagebox.showerror("参照エラー", f"ディレクトリの参照中にエラーが発生しました。\n\n詳細: {e}")

    def _select_output_file(self) -> None:
        """出力ファイルを選択"""
        try:
            current_path = self.output_file_var.get().strip()
            initial_file = "merged_output.pdf"  # デフォルト値を先に設定

            if current_path and os.path.dirname(current_path) and os.path.exists(os.path.dirname(current_path)):
                initial_dir = os.path.dirname(current_path)
                initial_file = os.path.basename(current_path)
            else:
                base_path = self.config.get('base_paths', 'google_drive')
                year = self.config.year
                education_base = self.config.get('directories', 'education_plan_base')
                initial_dir = os.path.join(base_path, year, education_base)
                config_file = self.config.get('output', 'merged_pdf')
                if config_file:
                    initial_file = config_file
                if not os.path.exists(initial_dir):
                    initial_dir = os.path.expanduser("~")

            file_path = filedialog.asksaveasfilename(
                title="出力ファイルを選択",
                initialdir=initial_dir,
                initialfile=initial_file,
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
            )
            if file_path:
                self.output_file_var.set(file_path)
                self.update_status(f"出力ファイルを選択: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("参照エラー", f"出力ファイルの参照中にエラーが発生しました。\n\n詳細: {e}")

    def _cancel_operation(self) -> None:
        """処理をキャンセル"""
        self._cancel_event.set()
        self.log("キャンセルリクエストを送信しました...", "warning")
        self.update_status("キャンセル処理中...")

    def _is_cancelled(self) -> bool:
        """キャンセル状態を返す（コールバック用、スレッドセーフ）"""
        return self._cancel_event.is_set()

    def _run_pdf_merge(self) -> None:
        """PDF統合を実行"""
        input_dir = self.input_dir_var.get()
        output_file = self.output_file_var.get()
        plan_type = self.plan_type_var.get()

        if not input_dir or not output_file:
            messagebox.showerror("入力エラー", "入力ディレクトリと出力ファイルの両方を指定してください。")
            return

        if not os.path.exists(input_dir):
            messagebox.showerror("パスエラー", f"入力ディレクトリが存在しません:\n\n{input_dir}")
            return

        # キャンセルフラグをリセット
        self._cancel_event.clear()

        def task():
            try:
                set_button_state(self.run_button, False, self.status_label, "🔄 実行中...")
                self.cancel_button.config(state="normal")
                self.progress.start(10)
                self.update_status("PDF統合を実行中...")

                self.log("=== PDF統合開始 ===", "info")
                self.log(f"入力: {input_dir}")
                self.log(f"出力: {output_file}")
                self.log(f"種別: {'教育計画' if plan_type == 'education' else '行事計画'}")

                # PDF統合処理を実行
                from pdf_converter import PDFConverter
                from pdf_processor import PDFProcessor
                from document_collector import DocumentCollector, PDFMergeOrchestrator, CancelledError

                temp_dir = self.config.get_temp_dir()
                ichitaro_settings = self.config.get('ichitaro')
                # キャンセルチェックコールバックを渡す
                converter = PDFConverter(temp_dir, ichitaro_settings, cancel_check=self._is_cancelled)
                processor = PDFProcessor(self.config)
                template_path = self.config.get_template_path()
                # キャンセルチェックコールバックを渡す
                collector = DocumentCollector(
                    converter, processor, template_path,
                    cancel_check=self._is_cancelled
                )
                orchestrator = PDFMergeOrchestrator(
                    self.config, converter, processor, collector,
                    cancel_check=self._is_cancelled
                )
                create_separators = (plan_type == "education")
                orchestrator.create_merged_pdf(input_dir, output_file, create_separators)

                self.log("=== PDF統合完了 ===", "success")
                set_button_state(self.run_button, True, self.status_label, "✅ 完了")
                self.update_status("PDF統合が完了しました")
                # スレッドセーフにダイアログを表示
                thread_safe_call(self.tab, lambda: messagebox.showinfo(
                    "✅ 完了", f"PDF統合が完了しました！\n\n出力ファイル:\n{output_file}"
                ))

            except CancelledError:
                self.log("=== キャンセルされました ===", "warning")
                set_button_state(self.run_button, True, self.status_label, "⚠️ キャンセル")
                self.update_status("PDF統合がキャンセルされました")
            except Exception as e:
                self.log(f"エラー: {e}", "error")
                set_button_state(self.run_button, True, self.status_label, "❌ エラー")
                self.update_status("PDF統合でエラーが発生しました")
                # スレッドセーフにダイアログを表示
                error_msg = str(e)
                thread_safe_call(self.tab, lambda: messagebox.showerror(
                    "❌ 実行エラー", f"PDF統合中にエラーが発生しました。\n\n詳細:\n{error_msg}"
                ))
            finally:
                def _cleanup():
                    try:
                        self.progress.stop()
                        self.cancel_button.config(state="disabled")
                    except Exception:
                        pass
                thread_safe_call(self.tab, _cleanup)

        thread = threading.Thread(target=task, daemon=True)
        thread.start()
