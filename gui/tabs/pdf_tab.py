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
from typing import Optional, TYPE_CHECKING

from gui.tabs.base_tab import BaseTab
from gui.utils import set_button_state, create_hover_button, thread_safe_call, open_file_or_folder, create_tooltip
from gui.ichitaro_dialog import IchitaroConversionDialog
from gui.styles import PADDING
from gui.ui_constants import (
    UIMessages, UILabels, UITooltips,
    UIWidgetSizes, UIIcons, UIColors
)
from pdf_converter import PDFConverter
from pdf_processor import PDFProcessor
from document_collector import DocumentCollector
from pdf_merge_orchestrator import PDFMergeOrchestrator
from exceptions import CancelledError
from path_validator import PathValidator

if TYPE_CHECKING:
    from config_loader import ConfigLoader

# ロガーの設定
logger = logging.getLogger(__name__)


class PDFTab(BaseTab):
    """PDF統合タブ"""

    def __init__(
        self,
        notebook: ttk.Notebook,
        config: "ConfigLoader",
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

        # 検証状態のラベル（後で作成）
        self.input_validation_label: Optional[tk.Label] = None
        self.output_validation_label: Optional[tk.Label] = None

        self._create_ui()
        self.add_to_notebook("📄 PDF統合")

        # 検証のデバウンス用タイマー
        self._validation_timer = None

        # 入力フィールドの変更を監視（デバウンス処理付き）
        self.input_dir_var.trace_add('write', lambda *args: self._schedule_validation())
        self.output_file_var.trace_add('write', lambda *args: self._schedule_validation())

        # 設定からデフォルトパスを読み込み
        self._load_default_paths()

    def _create_ui(self) -> None:
        """UIを構築"""
        # スクロール可能なメインコンテナ（BaseTabの共通メソッドを使用）
        self.canvas, _scrollbar, self.scrollable_frame = self.create_scrollable_container()

        # メインコンテナをスクロール可能フレーム内に配置
        main_container = self.scrollable_frame

        # 使い方ガイド（初心者向け）
        guide_frame = tk.LabelFrame(main_container, text="📖 使い方", font=("メイリオ", 10, "bold"))
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
        form_frame = tk.Frame(main_container)
        form_frame.pack(fill="x", padx=20, pady=15)

        LABEL_WIDTH = 18

        # 入力ディレクトリ選択
        tk.Label(form_frame, text="入力ディレクトリ:", width=LABEL_WIDTH, anchor="e").grid(row=0, column=0, sticky="e", padx=(15, 5), pady=6)

        input_entry = tk.Entry(form_frame, textvariable=self.input_dir_var, width=UIWidgetSizes.ENTRY_LARGE_WIDTH)
        input_entry.grid(row=0, column=1, padx=5, pady=6, sticky="ew")

        # プレースホルダー効果
        if not self.input_dir_var.get():
            input_entry.config(fg='gray')
            input_entry.insert(0, UILabels.PLACEHOLDER_DIR)
            input_entry.bind('<FocusIn>', lambda e: self._clear_placeholder(input_entry, UILabels.PLACEHOLDER_DIR))
            input_entry.bind('<FocusOut>', lambda e: self._restore_placeholder(input_entry, self.input_dir_var, UILabels.PLACEHOLDER_DIR))

        input_btn_frame = tk.Frame(form_frame)
        input_btn_frame.grid(row=0, column=2, padx=(5, 0), pady=6)

        def on_input_select_click():
            logger.info("入力ディレクトリ参照ボタンがクリックされました")
            self._select_input_dir()

        input_select_btn = tk.Button(input_btn_frame, text="📁", command=on_input_select_click, width=3)
        input_select_btn.pack(side="left", padx=1)
        create_tooltip(input_select_btn, UITooltips.TIP_FOLDER_BROWSE)

        input_open_btn = tk.Button(input_btn_frame, text="📂", command=self._open_input_dir, width=3)
        input_open_btn.pack(side="left", padx=1)
        create_tooltip(input_open_btn, UITooltips.TIP_FOLDER_OPEN)

        # 検証インジケーター
        self.input_validation_label = tk.Label(form_frame, text="", font=("メイリオ", 10), width=2)
        self.input_validation_label.grid(row=0, column=3, padx=(5, 15), pady=6)

        # 出力ファイル選択
        tk.Label(form_frame, text="出力ファイル:", width=LABEL_WIDTH, anchor="e").grid(row=1, column=0, sticky="e", padx=(15, 5), pady=6)

        output_entry = tk.Entry(form_frame, textvariable=self.output_file_var, width=UIWidgetSizes.ENTRY_LARGE_WIDTH)
        output_entry.grid(row=1, column=1, padx=5, pady=6, sticky="ew")

        output_btn_frame = tk.Frame(form_frame)
        output_btn_frame.grid(row=1, column=2, padx=(5, 0), pady=6)

        output_select_btn = tk.Button(output_btn_frame, text="💾", command=self._select_output_file, width=3)
        output_select_btn.pack(side="left", padx=1)
        create_tooltip(output_select_btn, UITooltips.TIP_FILE_BROWSE)

        output_open_btn = tk.Button(output_btn_frame, text="📂", command=self._open_output_dir, width=3)
        output_open_btn.pack(side="left", padx=1)
        create_tooltip(output_open_btn, UITooltips.TIP_FOLDER_OPEN)

        # 検証インジケーター
        self.output_validation_label = tk.Label(form_frame, text="", font=("メイリオ", 10), width=2)
        self.output_validation_label.grid(row=1, column=3, padx=(5, 15), pady=6)

        # 計画種別（自動判定結果の表示のみ）
        tk.Label(form_frame, text="計画種別:", width=LABEL_WIDTH, anchor="e").grid(row=2, column=0, sticky="e", padx=(15, 5), pady=6)
        self.plan_type_label = tk.Label(
            form_frame,
            text="自動判定中...",
            font=("メイリオ", 10),
            fg="#666",
            anchor="w"
        )
        self.plan_type_label.grid(row=2, column=1, sticky="w", padx=5, pady=6)
        create_tooltip(self.plan_type_label, "入力フォルダから自動判定されます")

        form_frame.columnconfigure(1, weight=1)

        # 実行ボタン
        button_frame = tk.Frame(main_container)
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
        self.status_label = tk.Label(main_container, text="", font=("メイリオ", 9), fg="gray")
        self.status_label.pack()

        # プログレスバー
        self.progress = ttk.Progressbar(main_container, mode='indeterminate')
        self.progress.pack(fill="x", padx=20, pady=5)

        # ログ表示
        self.create_log_frame(height=10, parent=main_container)
        # GUIログハンドラを設定（各モジュールのログをGUIに表示）
        self.setup_gui_logging()
        self.log("準備完了。入力ディレクトリと出力ファイルを選択して実行してください。", "info")

    def _select_input_dir(self) -> None:
        """入力ディレクトリを選択（pathlibベース）"""
        try:
            logger.info("ディレクトリ選択ダイアログを開きます")

            # tkinterの標準ダイアログを使用（sys.coinit_flagsでフリーズ解決済み）
            directory = filedialog.askdirectory(title="入力ディレクトリを選択")

            logger.info(f"ダイアログから戻りました: {directory if directory else 'キャンセル'}")

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
                    self._show_validation_error(error_msg)
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
            logger.info("出力ファイル選択ダイアログを開きます")

            # デフォルトの出力先をデスクトップに設定
            import os
            desktop_path = Path.home() / "Desktop"
            initial_dir = str(desktop_path) if desktop_path.exists() else str(Path.home())

            # tkinterの標準ダイアログを使用（sys.coinit_flagsでフリーズ解決済み）
            file_path = filedialog.asksaveasfilename(
                title="出力ファイルを選択",
                initialdir=initial_dir,
                initialfile="merged_output.pdf",
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
            )

            logger.info(f"ダイアログから戻りました: {file_path if file_path else 'キャンセル'}")

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
                    # 実行ボタンの状態を更新
                    self._update_run_button_state()
                else:
                    self._show_validation_error(error_msg)
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
        logger.info("PDF統合実行ボタンがクリックされました")

        # 入力値の取得
        input_dir_str = self.input_dir_var.get()
        output_file_str = self.output_file_var.get()
        plan_type = self.plan_type_var.get()

        logger.info(f"入力値: input_dir={input_dir_str}, output_file={output_file_str}, plan_type={plan_type}")

        # 空チェック
        if not input_dir_str or not output_file_str:
            logger.error(f"入力値が空です: input_dir={bool(input_dir_str)}, output_file={bool(output_file_str)}")
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
                from folder_structure_detector import FolderStructureDetector, PlanType

                detector = FolderStructureDetector()
                result = detector.detect_structure(str(directory_path))

                # UIスレッドで結果を反映
                def update_ui():
                    try:
                        if result.plan_type == PlanType.AMBIGUOUS:
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
            from folder_structure_detector import FolderStructureDetector, PlanType

            detector = FolderStructureDetector()
            result = detector.detect_structure(str(directory_path))

            if result.plan_type == PlanType.AMBIGUOUS:
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
        判定結果をラベルに表示

        Args:
            result: DetectionResult
        """
        plan_name = "教育計画" if result.plan_type.value == "education" else "行事計画"
        confidence_pct = int(result.confidence * 100)

        # アイコンを追加
        icon = "📚" if result.plan_type.value == "education" else "📅"

        # UIラベルを更新
        if hasattr(self, 'plan_type_label'):
            self.plan_type_label.config(
                text=f"{icon} {plan_name} (確信度: {confidence_pct}%)",
                fg="#2196F3" if confidence_pct >= 70 else "#FF9800"
            )

        # ステータスバーにも表示
        message = f"計画種別を自動判定: {plan_name} (確信度: {confidence_pct}%)"
        self.status_label.config(text=message, fg="green")
        self.log(message, "info")

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

    def _clear_placeholder(self, entry: tk.Entry, placeholder: str) -> None:
        """プレースホルダーをクリア"""
        if entry.get() == placeholder:
            entry.delete(0, tk.END)
            entry.config(fg='black')

    def _restore_placeholder(self, entry: tk.Entry, var: tk.StringVar, placeholder: str) -> None:
        """プレースホルダーを復元"""
        if not var.get():
            entry.config(fg='gray')
            entry.insert(0, placeholder)

    def _schedule_validation(self) -> None:
        """検証処理をスケジュール（デバウンス処理）"""
        # 既存のタイマーをキャンセル
        if self._validation_timer is not None:
            self.tab.after_cancel(self._validation_timer)

        # 300ms後に検証を実行（ユーザーの入力が落ち着いてから）
        self._validation_timer = self.tab.after(300, self._validate_inputs)

    def _validate_inputs(self) -> None:
        """入力フィールドの検証とビジュアルフィードバック"""
        # 入力ディレクトリの検証
        input_path = self.input_dir_var.get()
        if input_path and input_path != UILabels.PLACEHOLDER_DIR:
            is_valid, error_msg, validated_path = PathValidator.validate_directory(input_path, must_exist=True)
            if is_valid:
                self.input_validation_label.config(text=UIIcons.ICON_SUCCESS, fg=UIColors.VALID)
                create_tooltip(self.input_validation_label, "入力ディレクトリが存在します")
            else:
                self.input_validation_label.config(text=UIIcons.ICON_ERROR, fg=UIColors.INVALID)
                create_tooltip(self.input_validation_label, error_msg)
        else:
            self.input_validation_label.config(text="", fg='black')

        # 出力ファイルの検証
        output_path = self.output_file_var.get()
        if output_path and output_path != UILabels.PLACEHOLDER_FILE:
            is_valid, error_msg, validated_path = PathValidator.validate_file_path(output_path, must_exist=False, allowed_extensions=['.pdf'])
            if is_valid:
                self.output_validation_label.config(text=UIIcons.ICON_SUCCESS, fg=UIColors.VALID)
                create_tooltip(self.output_validation_label, "出力先のパスが有効です")
            else:
                self.output_validation_label.config(text=UIIcons.ICON_ERROR, fg=UIColors.INVALID)
                create_tooltip(self.output_validation_label, error_msg)
        else:
            self.output_validation_label.config(text="", fg='black')

        # 実行ボタンの有効/無効を更新
        self._update_run_button_state()

    def _update_run_button_state(self) -> None:
        """実行ボタンの状態を更新"""
        input_path = self.input_dir_var.get()
        output_path = self.output_file_var.get()

        logger.debug(f"実行ボタン状態チェック: input={input_path}, output={output_path}")

        # 両方が入力されており、プレースホルダーでない場合のみ有効
        if (input_path and input_path != UILabels.PLACEHOLDER_DIR and
            output_path and output_path != UILabels.PLACEHOLDER_FILE):
            # さらに実際にパスが有効かチェック
            input_valid, input_err, _ = PathValidator.validate_directory(input_path, must_exist=True)
            output_valid, output_err, _ = PathValidator.validate_file_path(output_path, must_exist=False, allowed_extensions=['.pdf'])

            logger.debug(f"パス検証結果: input_valid={input_valid}, output_valid={output_valid}")
            if not input_valid:
                logger.debug(f"入力パス検証エラー: {input_err}")
            if not output_valid:
                logger.debug(f"出力パス検証エラー: {output_err}")

            if input_valid and output_valid:
                logger.info("実行ボタンを有効化")
                self.run_button.config(state='normal')
            else:
                logger.warning(f"実行ボタンを無効化: input_valid={input_valid}, output_valid={output_valid}")
                self.run_button.config(state='disabled')
        else:
            logger.warning(f"実行ボタンを無効化: 入力が不十分 (input={bool(input_path)}, output={bool(output_path)})")
            self.run_button.config(state='disabled')

    def _load_default_paths(self) -> None:
        """設定からデフォルトパスを読み込む"""
        try:
            # 設定からGoogle Driveのベースパスを取得
            base_paths = self.config.get("base_paths") or {}
            google_drive_base = base_paths.get("google_drive", "")

            # 入力ディレクトリが未設定の場合、Google Driveのベースパスを設定
            if not self.input_dir_var.get() or self.input_dir_var.get() == UILabels.PLACEHOLDER_DIR:
                if google_drive_base:
                    # 教育計画のディレクトリパスを構築
                    year = self.config.get("year") or self.config.year or ""
                    year_short = self.config.get("year_short") or self.config.year_short or "R7"
                    directories = self.config.get("directories") or {}
                    education_plan_base = directories.get("education_plan_base", "")
                    education_plan = directories.get("education_plan", "")

                    if education_plan_base and education_plan and year:
                        # プレースホルダーを実際の値に置換
                        education_plan_base = education_plan_base.format(year_short=year_short)
                        education_plan = education_plan.format(year_short=year_short)

                        # フルパスを構築
                        default_input_path = Path(google_drive_base) / year / education_plan_base / education_plan

                        # パスが存在する場合のみ設定
                        if default_input_path.exists():
                            self.input_dir_var.set(str(default_input_path))
                            logger.info(f"デフォルト入力ディレクトリを設定: {default_input_path}")

            # 出力ファイルが未設定の場合、デスクトップのデフォルトファイル名を設定
            if not self.output_file_var.get() or self.output_file_var.get() == UILabels.PLACEHOLDER_FILE:
                desktop_path = Path.home() / "Desktop"
                output_config = self.config.get("output") or {}
                default_output_file = output_config.get("merged_pdf", "merged_output.pdf")

                if desktop_path.exists():
                    default_output_path = desktop_path / default_output_file
                    self.output_file_var.set(str(default_output_path))
                    logger.info(f"デフォルト出力ファイルを設定: {default_output_path}")

        except Exception as e:
            logger.warning(f"デフォルトパスの読み込みに失敗: {e}", exc_info=True)

    def _show_validation_error(self, error_msg: Optional[str]) -> None:
        """
        検証エラーを表示（共通メソッド）

        Args:
            error_msg: エラーメッセージ（Noneの場合は不明なエラー）
        """
        messagebox.showwarning(
            UIMessages.ERROR_VALIDATION,
            error_msg or UIMessages.ERROR_UNKNOWN
        )
