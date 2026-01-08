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
from gui.utils import set_button_state, create_hover_button, thread_safe_call, open_file_or_folder, create_tooltip
from gui.ichitaro_dialog import IchitaroConversionDialog
from gui.styles import PADDING
from gui.ui_constants import (
    UIMessages, UILabels, UITooltips, UIDialogTitles,
    UIWidgetSizes, UIIcons, UIColors
)
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
        # 使い方ガイド（初心者向け）
        guide_frame = tk.LabelFrame(self.tab, text="📖 使い方", font=("メイリオ", 10, "bold"))
        guide_frame.pack(fill="x", padx=PADDING['xlarge'], pady=(PADDING['large'], PADDING['medium']))

        guide_text = (
            "① PDFにしたいファイルが入っているフォルダを選択\n"
            "② 作成するPDFファイルの保存先と名前を決める\n"
            "③ 教育計画か行事計画を選ぶ\n"
            "④ 「PDF統合を実行」ボタンをクリック"
        )
        guide_label = tk.Label(
            guide_frame,
            text=guide_text,
            justify="left",
            font=("メイリオ", 9),
            fg="#333",
            padx=15,
            pady=10
        )
        guide_label.pack(anchor="w")

        # 入力フォームのフレーム
        form_frame = tk.Frame(self.tab)
        form_frame.pack(fill="x", padx=20, pady=15)

        LABEL_WIDTH = 18

        # 入力ディレクトリ選択
        tk.Label(form_frame, text="入力ディレクトリ:", width=LABEL_WIDTH, anchor="e").grid(row=0, column=0, sticky="e", padx=(15, 5), pady=6)
        tk.Entry(form_frame, textvariable=self.input_dir_var).grid(row=0, column=1, padx=5, pady=6, sticky="ew")

        input_btn_frame = tk.Frame(form_frame)
        input_btn_frame.grid(row=0, column=2, padx=(5, 15), pady=6)

        input_select_btn = tk.Button(input_btn_frame, text="📁", command=self._select_input_dir, width=3)
        input_select_btn.pack(side="left", padx=1)
        create_tooltip(input_select_btn, "フォルダを選択します")

        input_open_btn = tk.Button(input_btn_frame, text="📂", command=self._open_input_dir, width=3)
        input_open_btn.pack(side="left", padx=1)
        create_tooltip(input_open_btn, "選択したフォルダを開きます")

        # 出力ファイル選択
        tk.Label(form_frame, text="出力ファイル:", width=LABEL_WIDTH, anchor="e").grid(row=1, column=0, sticky="e", padx=(15, 5), pady=6)
        tk.Entry(form_frame, textvariable=self.output_file_var).grid(row=1, column=1, padx=5, pady=6, sticky="ew")

        output_btn_frame = tk.Frame(form_frame)
        output_btn_frame.grid(row=1, column=2, padx=(5, 15), pady=6)

        output_select_btn = tk.Button(output_btn_frame, text="💾", command=self._select_output_file, width=3)
        output_select_btn.pack(side="left", padx=1)
        create_tooltip(output_select_btn, "保存先とファイル名を指定します")

        output_open_btn = tk.Button(output_btn_frame, text="📂", command=self._open_output_dir, width=3)
        output_open_btn.pack(side="left", padx=1)
        create_tooltip(output_open_btn, "保存先フォルダを開きます")

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
            # initialdirの設定（フリーズ防止のため安全なパスのみ使用）
            initial_dir = None
            current_path = self.input_dir_var.get().strip()

            # 現在のパスが有効な場合はそれを使用
            if current_path:
                try:
                    current_dir = Path(current_path)
                    if current_dir.exists() and current_dir.is_dir():
                        initial_dir = str(current_dir)
                        logger.debug(f"現在のパスを使用: {initial_dir}")
                except Exception as e:
                    logger.warning(f"現在のパスの検証に失敗: {e}")

            # パスが無効な場合はinitialdirを指定しない（システムデフォルト）
            if not initial_dir:
                logger.debug("システムデフォルトのディレクトリから開始")

            # ファイルダイアログを表示
            if initial_dir:
                directory = filedialog.askdirectory(
                    title="入力ディレクトリを選択",
                    initialdir=initial_dir
                )
            else:
                directory = filedialog.askdirectory(
                    title="入力ディレクトリを選択"
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

                    # フォルダ構造の自動判定（バックグラウンドで実行してUIフリーズを防止）
                    self._detect_and_set_plan_type_async(validated_path)
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
            # initialdirの設定（フリーズ防止のため安全なパスのみ使用）
            initial_dir = None
            initial_file = "merged_output.pdf"
            current_path = self.output_file_var.get().strip()

            # 現在のパスが有効な場合はその親ディレクトリを使用
            if current_path:
                try:
                    current_file = Path(current_path)
                    parent_dir = current_file.parent
                    if parent_dir.exists() and parent_dir.is_dir():
                        initial_dir = str(parent_dir)
                        initial_file = current_file.name
                        logger.debug(f"現在のパスを使用: dir={initial_dir}, file={initial_file}")
                except Exception as e:
                    logger.warning(f"現在のパスの検証に失敗: {e}")

            # パスが無効な場合はinitialdirを指定しない（システムデフォルト）
            if not initial_dir:
                logger.debug("システムデフォルトのディレクトリから開始")

            # ファイルダイアログを表示
            if initial_dir:
                file_path = filedialog.asksaveasfilename(
                    title="出力ファイルを選択",
                    initialdir=initial_dir,
                    initialfile=initial_file,
                    defaultextension=".pdf",
                    filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
                )
            else:
                file_path = filedialog.asksaveasfilename(
                    title="出力ファイルを選択",
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

    def _open_input_dir(self) -> None:
        """入力ディレクトリをエクスプローラーで開く"""
        dir_path = self.input_dir_var.get().strip()
        if not dir_path:
            messagebox.showwarning(
                "フォルダが未選択",
                "まず「📁」ボタンでフォルダを選択してください。"
            )
            return

        def on_error(error_msg: str):
            messagebox.showerror(
                "フォルダが見つかりません",
                "指定されたフォルダが存在しません。\n\n"
                "「📁」ボタンをクリックして、正しいフォルダを選択し直してください。"
            )

        if open_file_or_folder(dir_path, on_error):
            self.update_status(f"フォルダを開きました: {Path(dir_path).name}")
            logger.info(f"入力ディレクトリを開きました: {dir_path}")

    def _open_output_dir(self) -> None:
        """出力ファイルの親フォルダをエクスプローラーで開く"""
        file_path = self.output_file_var.get().strip()
        if not file_path:
            messagebox.showwarning(
                "保存先が未設定",
                "まず「💾」ボタンで保存先を指定してください。"
            )
            return

        dir_path = str(Path(file_path).parent)

        def on_error(error_msg: str):
            messagebox.showerror(
                "フォルダが見つかりません",
                "保存先のフォルダが存在しません。\n\n"
                "「💾」ボタンをクリックして、正しい保存先を選択し直してください。"
            )

        if open_file_or_folder(dir_path, on_error):
            self.update_status(f"フォルダを開きました: {Path(dir_path).name}")
            logger.info(f"出力ディレクトリを開きました: {dir_path}")

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
                # GUI操作はすべてスレッドセーフに実行
                set_button_state(self.run_button, False, self.status_label, "🔄 実行中...")
                thread_safe_call(self.tab, lambda: self.cancel_button.config(state="normal"))
                thread_safe_call(self.tab, lambda: self.progress.start(10))
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
                    "完了", f"PDF統合が完了しました！\n\n出力ファイル:\n{output_file_path}"
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
                    "実行エラー", f"PDF統合中にエラーが発生しました。\n\n詳細:\n{error_msg}"
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

    def _detect_and_set_plan_type_async(self, directory_path: Path) -> None:
        """
        フォルダ構造を自動判定してplan_type_varを更新（非同期版・UIフリーズ防止）

        Args:
            directory_path: 判定対象のディレクトリPath
        """
        # ステータス更新
        self.update_status("フォルダ構造を自動判定中...")

        def task():
            try:
                from folder_structure_detector import FolderStructureDetector

                detector = FolderStructureDetector()
                result = detector.detect_structure(str(directory_path))

                # UIスレッドで結果を反映
                def update_ui():
                    try:
                        if result.plan_type == FolderStructureDetector.PlanType.AMBIGUOUS:
                            # 判定が曖昧な場合はダイアログで確認
                            self._show_plan_type_selection_dialog(result)
                        else:
                            # 確定判定の場合は自動設定
                            self.plan_type_var.set(result.plan_type.value)
                            self._update_plan_type_display(result)
                    except Exception as ui_error:
                        logger.error(f"UI更新エラー: {ui_error}", exc_info=True)

                self.tab.after(0, update_ui)

            except Exception as e:
                logger.error(f"フォルダ構造判定エラー: {e}", exc_info=True)
                # エラー時はデフォルト動作（手動選択のまま）
                def show_error():
                    self.update_status("フォルダ構造の自動判定をスキップしました")
                self.tab.after(0, show_error)

        # バックグラウンドスレッドで実行
        thread = threading.Thread(target=task, daemon=True, name="FolderStructureDetection")
        thread.start()

    def _detect_and_set_plan_type(self, directory_path: Path) -> None:
        """
        フォルダ構造を自動判定してplan_type_varを更新（同期版・後方互換性のため残す）

        Args:
            directory_path: 判定対象のディレクトリPath
        """
        try:
            from folder_structure_detector import FolderStructureDetector

            detector = FolderStructureDetector()
            result = detector.detect_structure(str(directory_path))

            if result.plan_type == FolderStructureDetector.PlanType.AMBIGUOUS:
                # 判定が曖昧な場合はダイアログで確認
                self._show_plan_type_selection_dialog(result)
            else:
                # 確定判定の場合は自動設定
                self.plan_type_var.set(result.plan_type.value)
                self._update_plan_type_display(result)

        except Exception as e:
            logger.error(f"フォルダ構造判定エラー: {e}", exc_info=True)
            # エラー時はデフォルト動作（手動選択のまま）

    def _update_plan_type_display(self, result) -> None:
        """
        判定結果をステータスラベルに表示

        Args:
            result: DetectionResult
        """
        plan_name = "教育計画" if result.plan_type.value == "education" else "行事計画"
        confidence_pct = int(result.confidence * 100)

        message = f"自動判定: {plan_name} (確信度: {confidence_pct}%)"
        self.status_label.config(text=message, fg="green")
        self.log(f"{message}", "info")

    def _show_plan_type_selection_dialog(self, result) -> None:
        """
        判定が曖昧な場合の選択ダイアログを表示

        Args:
            result: DetectionResult
        """
        from gui.plan_type_selection_dialog import PlanTypeSelectionDialog

        def on_selection(plan_type: str):
            """ダイアログでの選択結果を処理"""
            if plan_type:
                self.plan_type_var.set(plan_type)
                plan_name = "教育計画" if plan_type == "education" else "行事計画"
                self.update_status(f"計画種別を選択: {plan_name}")
                self.log(f"手動選択: {plan_name}", "info")

        dialog = PlanTypeSelectionDialog(self.tab, result, on_selection)
        self.tab.wait_window(dialog)
