"""
PDF統合タブ

PDF統合機能のUIを提供
2025年ベストプラクティス準拠版
"""
import logging
import tkinter as tk
from tkinter import ttk, messagebox
import threading
from pathlib import Path
from typing import Optional, TYPE_CHECKING

from gui.tabs.base_tab import BaseTab
from gui.utils import (
    set_button_state,
    create_hover_button,
    thread_safe_call,
    create_tooltip,
    attach_placeholder,
)
from gui.ichitaro_dialog import IchitaroConversionDialog
from gui.styles import COLORS, FONTS
from core.pdf_converter import PDFConverter
from core.pdf_processor import PDFProcessor
from core.document_collector import DocumentCollector
from core.pdf_merge_orchestrator import PDFMergeOrchestrator
from shared.exceptions import CancelledError
from shared.constants import PDFConstants
from infrastructure.path_validator import PathValidator

if TYPE_CHECKING:
    from infrastructure.config_loader import ConfigLoader

# ロガーの設定
logger = logging.getLogger(__name__)


_PLACEHOLDER_DIR = "フォルダを選択してください..."


class PDFTab(BaseTab):
    """PDF統合タブ"""

    def __init__(
        self,
        notebook: ttk.Notebook,
        config: "ConfigLoader",
        status_bar: tk.Label,
        input_dir_var: tk.StringVar,
        output_file_var: tk.StringVar,
        plan_type_var: tk.StringVar,
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
        self.add_to_notebook("PDF統合")

        # 検証のデバウンス用タイマー
        self._validation_timer = None

        # 入力フィールドの変更を監視（デバウンス処理付き）
        self.input_dir_var.trace_add("write", lambda *args: self._schedule_validation())
        self.output_file_var.trace_add(
            "write", lambda *args: self._schedule_validation()
        )

        # 設定からデフォルトパスを読み込み
        self._load_default_paths()

    def _create_ui(self) -> None:
        """UIを構築"""
        main_container = tk.Frame(self.tab)
        main_container.pack(fill="x")

        # 入力フォームのフレーム
        form_frame = tk.Frame(main_container)
        form_frame.pack(fill="x", padx=20, pady=15)

        LABEL_WIDTH = 18

        # 入力ディレクトリ選択
        tk.Label(
            form_frame, text="入力ディレクトリ:", width=LABEL_WIDTH, anchor="e"
        ).grid(row=0, column=0, sticky="e", padx=(15, 5), pady=6)

        input_entry = tk.Entry(form_frame, textvariable=self.input_dir_var, width=50)
        input_entry.grid(row=0, column=1, padx=5, pady=6, sticky="ew")
        attach_placeholder(input_entry, self.input_dir_var, _PLACEHOLDER_DIR)

        input_btn_frame = tk.Frame(form_frame)
        input_btn_frame.grid(row=0, column=2, padx=(5, 0), pady=6)

        def on_input_select_click():
            logger.debug("入力ディレクトリ参照ボタンがクリックされました")
            self._select_input_dir()

        input_select_btn = tk.Button(
            input_btn_frame, text="📁", command=on_input_select_click, width=3
        )
        input_select_btn.pack(side="left", padx=1)
        create_tooltip(input_select_btn, "フォルダ選択ダイアログを開きます")

        # 検証インジケーター
        self.input_validation_label = tk.Label(
            form_frame, text="", font=FONTS["default"], width=2
        )
        self.input_validation_label.grid(row=0, column=3, padx=(5, 15), pady=6)

        # 出力ファイル選択
        tk.Label(form_frame, text="出力ファイル:", width=LABEL_WIDTH, anchor="e").grid(
            row=1, column=0, sticky="e", padx=(15, 5), pady=6
        )

        output_entry = tk.Entry(form_frame, textvariable=self.output_file_var, width=50)
        output_entry.grid(row=1, column=1, padx=5, pady=6, sticky="ew")

        output_btn_frame = tk.Frame(form_frame)
        output_btn_frame.grid(row=1, column=2, padx=(5, 0), pady=6)

        output_select_btn = tk.Button(
            output_btn_frame, text="💾", command=self._select_output_file, width=3
        )
        output_select_btn.pack(side="left", padx=1)
        create_tooltip(output_select_btn, "ファイル選択ダイアログを開きます")

        # 検証インジケーター
        self.output_validation_label = tk.Label(
            form_frame, text="", font=FONTS["default"], width=2
        )
        self.output_validation_label.grid(row=1, column=3, padx=(5, 15), pady=6)

        # 計画種別（自動判定結果の表示のみ）
        tk.Label(form_frame, text="計画種別:", width=LABEL_WIDTH, anchor="e").grid(
            row=2, column=0, sticky="e", padx=(15, 5), pady=6
        )
        self.plan_type_label = tk.Label(
            form_frame,
            text="自動判定中...",
            font=FONTS["default"],
            fg=COLORS["text_secondary"],
            anchor="w",
        )
        self.plan_type_label.grid(row=2, column=1, sticky="w", padx=5, pady=6)
        create_tooltip(self.plan_type_label, "入力フォルダから自動判定されます")

        # PDF圧縮オプション
        tk.Label(form_frame, text="オプション:", width=LABEL_WIDTH, anchor="e").grid(
            row=3, column=0, sticky="e", padx=(15, 5), pady=6
        )
        self.compress_var = tk.BooleanVar(value=True)
        compress_check = tk.Checkbutton(
            form_frame,
            text="Ghostscriptで圧縮する",
            variable=self.compress_var,
            font=FONTS["default"],
        )
        compress_check.grid(row=3, column=1, sticky="w", padx=5, pady=6)
        create_tooltip(
            compress_check, "Ghostscriptが利用可能な場合、最終PDFを圧縮します"
        )

        form_frame.columnconfigure(1, weight=1)

        # 実行ボタン
        button_frame = tk.Frame(main_container)
        button_frame.pack(pady=15)

        self.run_button = create_hover_button(
            button_frame,
            text="▶ PDF統合を実行",
            command=self._run_pdf_merge,
            color="primary",
            font=FONTS["subheading"],
            width=28,
            height=2,
        )
        self.run_button.pack(side="left", padx=5)

        self.cancel_button = tk.Button(
            button_frame,
            text="✕ キャンセル",
            command=self._cancel_operation,
            font=FONTS["default"],
            bg=COLORS["error"],
            fg="white",
            width=12,
            height=2,
            state="disabled",
        )
        self.cancel_button.pack(side="left", padx=5)

        # ステータスラベル
        self.status_label = tk.Label(
            main_container, text="", font=FONTS["small"], fg="gray"
        )
        self.status_label.pack()

        # プログレスバー
        self.progress = ttk.Progressbar(
            main_container,
            mode="determinate",
            maximum=PDFConstants.MERGE_STEPS_WITH_COMPRESS,
        )
        self.progress.pack(fill="x", padx=20, pady=5)

        # ログ表示（タブ直下に配置し、ウィンドウリサイズに追従させる）
        self.create_log_frame(height=8)
        # GUIログハンドラを設定（各モジュールのログをGUIに表示）
        self.setup_gui_logging()
        self.log(
            "準備完了。入力ディレクトリと出力ファイルを選択して実行してください。",
            "info",
        )

    def _select_input_dir(self) -> None:
        """入力ディレクトリを選択"""
        validated_path = self.ask_folder(title="入力ディレクトリを選択")
        if validated_path:
            self.input_dir_var.set(str(validated_path))
            self.update_status(f"入力ディレクトリを選択: {validated_path.name}")
            logger.info(f"入力ディレクトリを選択: {validated_path}")
            # フォルダ構造の自動判定（バックグラウンド実行でUIフリーズ防止）
            self._detect_and_set_plan_type_async(validated_path)

    def _select_output_file(self) -> None:
        """出力ファイルを選択"""
        desktop_path = Path.home() / "Desktop"
        initial_dir = str(desktop_path) if desktop_path.exists() else str(Path.home())
        initial_file = (
            Path(self.output_file_var.get()).name if self.output_file_var.get() else ""
        )
        validated_path = self.ask_file_save(
            title="出力ファイルを選択",
            initial_dir=initial_dir,
            initial_file=initial_file,
            default_extension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            allowed_extensions=[".pdf"],
        )
        if validated_path:
            self.output_file_var.set(str(validated_path))
            self.update_status(f"出力ファイルを選択: {validated_path.name}")
            logger.info(f"出力ファイルを選択: {validated_path}")
            self._validate_inputs()

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
        logger.debug("PDF統合実行ボタンがクリックされました")

        # 入力値の取得
        input_dir_str = self.input_dir_var.get()
        output_file_str = self.output_file_var.get()
        plan_type = self.plan_type_var.get()

        logger.debug(
            f"入力値: input_dir={input_dir_str}, output_file={output_file_str}, plan_type={plan_type}"
        )

        # 空チェック
        if not input_dir_str or not output_file_str:
            logger.error(
                f"入力値が空です: input_dir={bool(input_dir_str)}, output_file={bool(output_file_str)}"
            )
            messagebox.showerror(
                "入力エラー", "入力ディレクトリと出力ファイルの両方を指定してください。"
            )
            return

        # 入力ディレクトリの検証
        input_dir_path = self.validate_path(input_dir_str, "directory")
        if not input_dir_path:
            return

        # 出力ファイルの検証
        output_file_path = self.validate_path(
            output_file_str, "file", must_exist=False, allowed_extensions=[".pdf"]
        )
        if not output_file_path:
            return

        logger.debug(f"パス検証完了 - 入力: {input_dir_path}, 出力: {output_file_path}")

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
                                self.tab, cancel_callback=self._cancel_operation
                            )
                        ichitaro_dialog.update_message(message)
                    else:
                        if ichitaro_dialog:
                            ichitaro_dialog.close()
                            ichitaro_dialog = None

                thread_safe_call(self.tab, _handle)

            try:
                # GUI操作はすべてスレッドセーフに実行
                set_button_state(
                    self.run_button, False, self.status_label, "🔄 実行中..."
                )
                thread_safe_call(
                    self.tab, lambda: self.cancel_button.config(state="normal")
                )
                thread_safe_call(self.tab, lambda: self.progress.configure(value=0))
                self.update_status("PDF統合を実行中...")

                self.log("=== PDF統合開始 ===", "info")
                self.log(f"入力: {input_dir_path}")
                self.log(f"出力: {output_file_path}")
                self.log(
                    f"種別: {'教育計画' if plan_type == 'education' else '行事計画'}"
                )

                # PDF統合処理を実行（Pathオブジェクトを文字列に変換）
                input_dir_str_final = str(input_dir_path)
                output_file_str_final = str(output_file_path)

                temp_dir = self.config.get_temp_dir(cleanup_old=True)
                ichitaro_settings = self.config.get("ichitaro")

                converter = PDFConverter(
                    temp_dir,
                    ichitaro_settings,
                    cancel_check=self._is_cancelled,
                    dialog_callback=dialog_callback,
                    config=self.config,
                )
                processor = PDFProcessor(self.config)
                collector = DocumentCollector(
                    converter, processor, cancel_check=self._is_cancelled
                )

                use_compress = self.compress_var.get()
                total_steps = (
                    PDFConstants.MERGE_STEPS_WITH_COMPRESS
                    if use_compress
                    else PDFConstants.MERGE_STEPS
                )

                def on_progress(step: int, total: int, message: str) -> None:
                    thread_safe_call(
                        self.tab, lambda: self.progress.configure(value=step)
                    )

                thread_safe_call(
                    self.tab, lambda: self.progress.configure(maximum=total_steps)
                )

                orchestrator = PDFMergeOrchestrator(
                    self.config,
                    converter,
                    processor,
                    collector,
                    cancel_check=self._is_cancelled,
                    progress_callback=on_progress,
                )
                create_separators = plan_type == "education"
                orchestrator.create_merged_pdf(
                    input_dir_str_final,
                    output_file_str_final,
                    create_separators,
                    compress=use_compress,
                )

                thread_safe_call(
                    self.tab, lambda: self.progress.configure(value=total_steps)
                )
                self.log("=== PDF統合完了 ===", "success")
                set_button_state(self.run_button, True, self.status_label, "✅ 完了")
                self.update_status("PDF統合が完了しました")
                thread_safe_call(
                    self.tab,
                    lambda: messagebox.showinfo(
                        "完了",
                        f"PDF統合が完了しました！\n\n出力ファイル:\n{output_file_path}",
                    ),
                )

            except CancelledError:
                self.log("=== キャンセルされました ===", "warning")
                set_button_state(
                    self.run_button, True, self.status_label, "⚠️ キャンセル"
                )
                self.update_status("PDF統合がキャンセルされました")
            except Exception as e:
                self.log(f"エラー: {e}", "error")
                set_button_state(self.run_button, True, self.status_label, "❌ エラー")
                self.update_status("PDF統合でエラーが発生しました")
                # スレッドセーフにダイアログを表示
                error_msg = str(e)
                thread_safe_call(
                    self.tab,
                    lambda: messagebox.showerror(
                        "実行エラー",
                        f"PDF統合中にエラーが発生しました。\n\n詳細:\n{error_msg}",
                    ),
                )
            finally:

                def _cleanup():
                    try:
                        self.progress.configure(value=0)
                        self.cancel_button.config(state="disabled")
                    except Exception:
                        pass

                thread_safe_call(self.tab, _cleanup)

                # ダイアログが残っていたら閉じる（変数キャプチャのTOCTOU回避）
                dialog_ref = ichitaro_dialog
                if dialog_ref:
                    thread_safe_call(self.tab, lambda: dialog_ref.close())

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
                from core.folder_structure_detector import (
                    FolderStructureDetector,
                    PlanType,
                )

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
                            self._update_output_filename(result.plan_type.value)
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
        thread = threading.Thread(
            target=task, daemon=True, name="FolderStructureDetection"
        )
        thread.start()

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
        if hasattr(self, "plan_type_label"):
            self.plan_type_label.config(
                text=f"{icon} {plan_name} (確信度: {confidence_pct}%)",
                fg=COLORS["primary"]
                if confidence_pct >= 70
                else COLORS["warning_mild"],
            )

        # ステータスバーにも表示
        message = f"計画種別を自動判定: {plan_name} (確信度: {confidence_pct}%)"
        self.status_label.config(text=message, fg="green")
        self.log(message, "info")

    def _update_output_filename(self, plan_type: str) -> None:
        """
        計画タイプに応じて出力ファイル名を自動更新

        ユーザーが手動で設定したカスタムファイル名は上書きしない。
        自動生成パターン（年度_計画名.pdf）またはデフォルト名の場合のみ更新。

        Args:
            plan_type: "education" or "event"
        """
        year_short = self.config.year_short
        plan_name = "教育計画" if plan_type == "education" else "行事計画"
        new_filename = f"{year_short}_{plan_name}.pdf"

        current_output = self.output_file_var.get().strip()

        # ユーザーがカスタマイズしたファイル名は上書きしない
        if current_output:
            current_name = Path(current_output).name
            auto_patterns = {
                f"{year_short}_教育計画.pdf",
                f"{year_short}_行事計画.pdf",
            }
            if current_name not in auto_patterns:
                logger.info(f"カスタムファイル名のため更新スキップ: {current_name}")
                return
            output_dir = Path(current_output).parent
        else:
            desktop = Path.home() / "Desktop"
            output_dir = desktop if desktop.exists() else Path.home()

        new_path = output_dir / new_filename
        self.output_file_var.set(str(new_path))
        logger.info(f"出力ファイル名を更新: {new_path}")

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
                self._update_output_filename(plan_type)

        dialog = PlanTypeSelectionDialog(self.tab, result, on_selection)
        self.tab.wait_window(dialog)

    def _schedule_validation(self) -> None:
        """検証処理をスケジュール（デバウンス処理）"""
        # 既存のタイマーをキャンセル
        if self._validation_timer is not None:
            self.tab.after_cancel(self._validation_timer)

        # 300ms後に検証を実行（ユーザーの入力が落ち着いてから）
        self._validation_timer = self.tab.after(300, self._validate_inputs)

    def _validate_inputs(self) -> None:
        """入力フィールドの検証、ビジュアルフィードバック、実行ボタン状態の更新"""
        input_path = self.input_dir_var.get()
        output_path = self.output_file_var.get()

        # 入力ディレクトリの検証
        input_valid = False
        if input_path and input_path != _PLACEHOLDER_DIR:
            input_valid, error_msg, _ = PathValidator.validate_directory(
                input_path, must_exist=True
            )
            if input_valid:
                self.input_validation_label.config(text="✓", fg=COLORS["valid"])
            else:
                self.input_validation_label.config(text="✗", fg=COLORS["invalid"])
        else:
            self.input_validation_label.config(text="", fg="black")

        # 出力ファイルの検証
        output_valid = False
        if output_path:
            output_valid, error_msg, _ = PathValidator.validate_file_path(
                output_path, must_exist=False, allowed_extensions=[".pdf"]
            )
            if output_valid:
                self.output_validation_label.config(text="✓", fg=COLORS["valid"])
            else:
                self.output_validation_label.config(text="✗", fg=COLORS["invalid"])
        else:
            self.output_validation_label.config(text="", fg="black")

        # 実行ボタンの状態を更新
        if input_valid and output_valid:
            self.run_button.config(state="normal")
        else:
            self.run_button.config(state="disabled")

    def _load_default_paths(self) -> None:
        """設定からデフォルトパスを読み込む"""
        try:
            # 設定からGoogle Driveのベースパスを取得
            base_paths = self.config.get("base_paths") or {}
            google_drive_base = base_paths.get("google_drive", "")

            # 入力ディレクトリが未設定の場合、Google Driveのベースパスを設定
            if (
                not self.input_dir_var.get()
                or self.input_dir_var.get() == _PLACEHOLDER_DIR
            ):
                if google_drive_base:
                    # 教育計画のディレクトリパスを構築
                    year = self.config.get("year") or self.config.year or ""
                    year_short = (
                        self.config.get("year_short") or self.config.year_short or "R7"
                    )
                    directories = self.config.get("directories") or {}
                    education_plan_base = directories.get("education_plan_base", "")
                    education_plan = directories.get("education_plan", "")

                    if education_plan_base and education_plan and year:
                        # プレースホルダーを実際の値に置換
                        education_plan_base = education_plan_base.format(
                            year_short=year_short
                        )
                        education_plan = education_plan.format(year_short=year_short)

                        # フルパスを構築
                        default_input_path = (
                            Path(google_drive_base)
                            / year
                            / education_plan_base
                            / education_plan
                        )

                        # パスが存在する場合のみ設定
                        if default_input_path.exists():
                            self.input_dir_var.set(str(default_input_path))
                            logger.info(
                                f"デフォルト入力ディレクトリを設定: {default_input_path}"
                            )

            # 出力ファイルは入力ディレクトリ選択時に自動設定されるため、
            # ここでは設定しない（merged_output.pdfの無意味なデフォルト値を排除）

        except Exception as e:
            logger.warning(f"デフォルトパスの読み込みに失敗: {e}", exc_info=True)
