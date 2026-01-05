"""
PDF統合タブ

PDF統合機能のUIを提供
2025年ベストプラクティス準拠版
"""
import logging
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
from pathlib import Path
from typing import Any, Optional

from gui.tabs.base_tab import BaseTab
from gui.utils import set_button_state, create_hover_button, thread_safe_call
from gui.ichitaro_dialog import IchitaroConversionDialog
from pdf_converter import PDFConverter
from pdf_processor import PDFProcessor
from document_collector import DocumentCollector, PDFMergeOrchestrator, CancelledError
from path_validator import PathValidator

# ロガーの設定
logger = logging.getLogger(__name__)


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
        """入力ディレクトリを選択（pathlibベース）"""
        try:
            current_path_str = self.input_dir_var.get().strip()

            # 設定からフォールバックディレクトリを取得
            fallback = None
            try:
                default_input = self.config.get_education_plan_path()
                if default_input:
                    fallback = Path(default_input)
            except Exception:
                pass

            # 安全な初期ディレクトリを取得
            initial_dir = PathValidator.get_safe_initial_dir(current_path_str, fallback)

            logger.debug(f"ファイルダイアログを開きます: initial_dir={initial_dir}")

            # ファイルダイアログを表示
            directory = filedialog.askdirectory(
                title="入力ディレクトリを選択",
                initialdir=str(initial_dir)
            )

            if directory:
                # 選択されたパスを検証
                is_valid, error_msg, validated_path = PathValidator.validate_directory(
                    directory,
                    must_exist=True
                )

                if is_valid and validated_path:
                    self.input_dir_var.set(str(validated_path))
                    self.update_status(f"入力ディレクトリを選択: {validated_path.name}")
                    logger.info(f"入力ディレクトリを選択: {validated_path}")
                else:
                    messagebox.showwarning("検証エラー", error_msg or "不明なエラー")
            else:
                logger.debug("ディレクトリ選択がキャンセルされました")

        except Exception as e:
            logger.error(f"ディレクトリ選択エラー: {e}", exc_info=True)
            messagebox.showerror(
                "参照エラー",
                f"ディレクトリの参照中にエラーが発生しました。\n\n詳細: {e}"
            )

    def _select_output_file(self) -> None:
        """出力ファイルを選択（pathlibベース）"""
        try:
            current_path_str = self.output_file_var.get().strip()
            initial_file = "merged_output.pdf"

            # 現在のパスから初期情報を取得
            initial_dir = None
            if current_path_str:
                try:
                    current_path = Path(current_path_str)
                    if current_path.parent.exists():
                        initial_dir = current_path.parent
                        initial_file = current_path.name
                except:
                    pass

            # フォールバック: 設定から取得
            if not initial_dir:
                try:
                    base_path = self.config.get('base_paths', 'google_drive')
                    year = self.config.year
                    education_base = self.config.get('directories', 'education_plan_base')
                    config_dir = Path(base_path) / year / education_base
                    if config_dir.exists():
                        initial_dir = config_dir

                    config_file = self.config.get('output', 'merged_pdf')
                    if config_file:
                        initial_file = config_file
                except Exception:
                    pass

            # 最終フォールバック: ホームディレクトリ
            if not initial_dir:
                initial_dir = Path.home()

            logger.debug(f"ファイルダイアログを開きます: initial_dir={initial_dir}, initial_file={initial_file}")

            # ファイルダイアログを表示
            file_path = filedialog.asksaveasfilename(
                title="出力ファイルを選択",
                initialdir=str(initial_dir),
                initialfile=initial_file,
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
            )

            if file_path:
                # 選択されたパスを検証
                is_valid, error_msg, validated_path = PathValidator.validate_file_path(
                    file_path,
                    must_exist=False,
                    allowed_extensions=['.pdf']
                )

                if is_valid and validated_path:
                    self.output_file_var.set(str(validated_path))
                    self.update_status(f"出力ファイルを選択: {validated_path.name}")
                    logger.info(f"出力ファイルを選択: {validated_path}")
                else:
                    messagebox.showwarning("検証エラー", error_msg or "不明なエラー")
            else:
                logger.debug("出力ファイル選択がキャンセルされました")

        except Exception as e:
            logger.error(f"出力ファイル選択エラー: {e}", exc_info=True)
            messagebox.showerror(
                "参照エラー",
                f"出力ファイルの参照中にエラーが発生しました。\n\n詳細: {e}"
            )

    def _cancel_operation(self) -> None:
        """処理をキャンセル"""
        self._cancel_event.set()
        self.log("キャンセルリクエストを送信しました...", "warning")
        self.update_status("キャンセル処理中...")

    def _is_cancelled(self) -> bool:
        """キャンセル状態を返す（コールバック用、スレッドセーフ）"""
        return self._cancel_event.is_set()

    def _run_pdf_merge(self) -> None:
        """PDF統合を実行（pathlibベース、2025年ベストプラクティス準拠）"""
        # 入力値の取得
        input_dir_str = self.input_dir_var.get()
        output_file_str = self.output_file_var.get()
        plan_type = self.plan_type_var.get()

        # 空チェック
        if not input_dir_str or not output_file_str:
            messagebox.showerror("入力エラー", "入力ディレクトリと出力ファイルの両方を指定してください。")
            return

        # 入力ディレクトリの検証
        is_valid_dir, error_msg_dir, input_dir_path = PathValidator.validate_directory(
            input_dir_str,
            must_exist=True
        )

        if not is_valid_dir or not input_dir_path:
            logger.error(f"入力ディレクトリの検証エラー: {error_msg_dir}")
            messagebox.showerror("パスエラー", error_msg_dir or "入力ディレクトリが無効です")
            return

        # 出力ファイルの検証
        is_valid_file, error_msg_file, output_file_path = PathValidator.validate_file_path(
            output_file_str,
            must_exist=False,
            allowed_extensions=['.pdf']
        )

        if not is_valid_file or not output_file_path:
            logger.error(f"出力ファイルの検証エラー: {error_msg_file}")
            messagebox.showerror("パスエラー", error_msg_file or "出力ファイルパスが無効です")
            return

        logger.info(f"パス検証完了 - 入力: {input_dir_path}, 出力: {output_file_path}")

        # キャンセルフラグをリセット
        self._cancel_event.clear()

        def task():
            ichitaro_dialog = None

            def dialog_callback(message: str, show: bool):
                """一太郎変換ダイアログの表示/非表示"""
                nonlocal ichitaro_dialog

                def _handle():
                    nonlocal ichitaro_dialog
                    if show:
                        if not ichitaro_dialog:
                            ichitaro_dialog = IchitaroConversionDialog(
                                self.tab,
                                cancel_callback=self._cancel_operation
                            )
                        ichitaro_dialog.update_message(message)
                    else:
                        if ichitaro_dialog:
                            ichitaro_dialog.close()
                            ichitaro_dialog = None

                thread_safe_call(self.tab, _handle)

            try:
                set_button_state(self.run_button, False, self.status_label, "🔄 実行中...")
                self.cancel_button.config(state="normal")
                self.progress.start(10)
                self.update_status("PDF統合を実行中...")

                self.log("=== PDF統合開始 ===", "info")
                self.log(f"入力: {input_dir_path}")
                self.log(f"出力: {output_file_path}")
                self.log(f"種別: {'教育計画' if plan_type == 'education' else '行事計画'}")

                # PDF統合処理を実行（Pathオブジェクトを文字列に変換）
                input_dir_str_final = str(input_dir_path)
                output_file_str_final = str(output_file_path)

                self.log("一時ディレクトリを取得中...", "info")
                temp_dir = self.config.get_temp_dir()
                self.log(f"一時ディレクトリ: {temp_dir}", "info")

                self.log("設定を読み込み中...", "info")
                ichitaro_settings = self.config.get('ichitaro')

                self.log("PDFコンバーターを初期化中...", "info")
                converter = PDFConverter(
                    temp_dir,
                    ichitaro_settings,
                    cancel_check=self._is_cancelled,
                    dialog_callback=dialog_callback,
                    config=self.config
                )

                self.log("PDFプロセッサーを初期化中...", "info")
                processor = PDFProcessor(self.config)

                self.log("ドキュメントコレクターを初期化中...", "info")
                collector = DocumentCollector(
                    converter, processor,
                    cancel_check=self._is_cancelled
                )

                self.log("オーケストレーターを初期化中...", "info")
                orchestrator = PDFMergeOrchestrator(
                    self.config, converter, processor, collector,
                    cancel_check=self._is_cancelled
                )

                self.log("PDF統合処理を開始します...", "info")
                create_separators = (plan_type == "education")
                orchestrator.create_merged_pdf(input_dir_str_final, output_file_str_final, create_separators)

                self.log("=== PDF統合完了 ===", "success")
                set_button_state(self.run_button, True, self.status_label, "✅ 完了")
                self.update_status("PDF統合が完了しました")
                thread_safe_call(self.tab, lambda: messagebox.showinfo(
                    "✅ 完了", f"PDF統合が完了しました！\n\n出力ファイル:\n{output_file_path}"
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

                # ダイアログが残っていたら閉じる
                if ichitaro_dialog:
                    thread_safe_call(self.tab, lambda: ichitaro_dialog.close())

        thread = threading.Thread(target=task, daemon=True)
        thread.start()
