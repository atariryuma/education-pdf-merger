"""
Excel処理タブ

Excel自動更新機能のUIを提供
"""
import logging
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
from typing import Any, Tuple, TYPE_CHECKING, Optional
from pathlib import Path

try:
    import win32com.client
    import pythoncom
except ImportError:
    win32com = None
    pythoncom = None

from gui.tabs.base_tab import BaseTab
from gui.utils import set_button_state, create_hover_button, open_file_or_folder, create_tooltip, thread_safe_call
from path_validator import PathValidator
from update_excel_files import ExcelTransfer

if TYPE_CHECKING:
    from config_loader import ConfigLoader

# ロガーの設定
logger = logging.getLogger(__name__)


class ExcelTab(BaseTab):
    """Excel処理タブ"""

    def __init__(self, notebook: ttk.Notebook, config: "ConfigLoader", status_bar: tk.Label) -> None:
        """
        Excel処理タブの初期化

        Args:
            notebook: タブを追加するNotebookウィジェット
            config: ConfigLoaderインスタンス
            status_bar: ステータスバーのLabelウィジェット
        """
        super().__init__(notebook, config, status_bar)

        # セッション内でのファイルパス管理（config.jsonには保存しない）
        self.ref_file_path: Optional[str] = None
        self.target_file_path: Optional[str] = None

        # 設定タブへの参照（後から設定される）
        self.settings_tab: Optional[Any] = None

        self._create_ui()
        self.add_to_notebook("📊 Excel処理")

    def set_settings_tab(self, settings_tab: Any) -> None:
        """設定タブへの参照を設定（app.pyから呼ばれる）"""
        self.settings_tab = settings_tab

    def _create_ui(self) -> None:
        """UIを構築"""
        # スクロール可能なメインコンテナ（BaseTabの共通メソッドを使用）
        self.canvas, _scrollbar, self.scrollable_frame = self.create_scrollable_container()

        # メインコンテナをスクロール可能フレーム内に配置
        main_container = self.scrollable_frame

        # 説明フレーム
        info_frame = tk.LabelFrame(main_container, text="📋 Excel自動更新機能", font=("メイリオ", 11, "bold"))
        info_frame.pack(fill="x", padx=20, pady=15)

        info_text = "年間行事計画（編集用）から様式4へ自動的にデータを転記します。"
        tk.Label(info_frame, text=info_text, justify="left", font=("メイリオ", 10)).pack(pady=(15, 5), padx=15)

        steps_text = (
            "📝 実行手順\n\n"
            "1️⃣ 「🔍 自動検出」でExcelで開いているファイルを検出\n"
            "   または「📁 ファイルを選択」で手動選択\n\n"
            "2️⃣ ファイルパスが表示され、●マークが緑色になったことを確認\n\n"
            "3️⃣ 「Excelデータ更新を実行」ボタンをクリック\n\n"
            "4️⃣ 処理完了後、内容を確認してExcelで保存"
        )
        tk.Label(info_frame, text=steps_text, justify="left", font=("メイリオ", 9), fg="#333").pack(pady=(5, 15), padx=15, anchor="w")

        # ファイル選択フレーム
        file_frame = tk.LabelFrame(main_container, text="📂 対象ファイル", font=("メイリオ", 11, "bold"))
        file_frame.pack(fill="x", padx=20, pady=10)

        # 自動検出ボタン
        detect_frame = tk.Frame(file_frame)
        detect_frame.pack(pady=10)

        detect_btn = create_hover_button(
            detect_frame,
            text="🔍 開いているExcelファイルを自動検出",
            command=self._auto_detect_files,
            color="success",
            font=("メイリオ", 10, "bold"),
            width=36,
            height=1
        )
        detect_btn.pack()
        create_tooltip(detect_btn, "Excelで開いている「編集用」と「様式4」を自動検出します")

        # セパレーター
        ttk.Separator(file_frame, orient="horizontal").pack(fill="x", padx=15, pady=10)

        # 参照元ファイル
        ref_frame = tk.Frame(file_frame)
        ref_frame.pack(fill="x", padx=15, pady=8)

        tk.Label(ref_frame, text="参照元:", width=10, anchor="w", font=("メイリオ", 9, "bold")).pack(side="left")

        self.ref_status = tk.Label(ref_frame, text="●", fg="gray", font=("メイリオ", 12))
        self.ref_status.pack(side="left", padx=(0, 5))

        self.ref_label = tk.Label(
            ref_frame,
            text="未選択",
            font=("メイリオ", 9),
            fg="#666",
            anchor="w",
            cursor="hand2"
        )
        self.ref_label.pack(side="left", fill="x", expand=True, padx=5)
        self.ref_label.bind("<Button-1>", lambda e: self._open_selected_file(self.ref_file_path))
        create_tooltip(self.ref_label, "クリックでExcelファイルを開きます")

        ref_btn_frame = tk.Frame(ref_frame)
        ref_btn_frame.pack(side="right")

        tk.Button(
            ref_btn_frame,
            text="📁 ファイルを選択",
            command=lambda: self._select_file("reference"),
            font=("メイリオ", 8),
            width=14
        ).pack(side="left", padx=2)

        tk.Button(
            ref_btn_frame,
            text="📂",
            command=lambda: self._open_selected_file(self.ref_file_path),
            width=3,
            font=("メイリオ", 9)
        ).pack(side="left", padx=2)

        # 対象ファイル
        target_frame = tk.Frame(file_frame)
        target_frame.pack(fill="x", padx=15, pady=8)

        tk.Label(target_frame, text="対象:", width=10, anchor="w", font=("メイリオ", 9, "bold")).pack(side="left")

        self.target_status = tk.Label(target_frame, text="●", fg="gray", font=("メイリオ", 12))
        self.target_status.pack(side="left", padx=(0, 5))

        self.target_label = tk.Label(
            target_frame,
            text="未選択",
            font=("メイリオ", 9),
            fg="#666",
            anchor="w",
            cursor="hand2"
        )
        self.target_label.pack(side="left", fill="x", expand=True, padx=5)
        self.target_label.bind("<Button-1>", lambda e: self._open_selected_file(self.target_file_path))
        create_tooltip(self.target_label, "クリックでExcelファイルを開きます")

        target_btn_frame = tk.Frame(target_frame)
        target_btn_frame.pack(side="right")

        tk.Button(
            target_btn_frame,
            text="📁 ファイルを選択",
            command=lambda: self._select_file("target"),
            font=("メイリオ", 8),
            width=14
        ).pack(side="left", padx=2)

        tk.Button(
            target_btn_frame,
            text="📂",
            command=lambda: self._open_selected_file(self.target_file_path),
            width=3,
            font=("メイリオ", 9)
        ).pack(side="left", padx=2)

        # セパレーター（視覚的な区切り）
        ttk.Separator(main_container, orient="horizontal").pack(fill="x", padx=20, pady=10)

        # 行事名初期設定エリア（年1回の使用を想定）
        event_setup_frame = tk.LabelFrame(
            main_container,
            text="📥 行事名の初期設定（年1回・初回セットアップ用）",
            font=("メイリオ", 10, "bold"),
            bg="#F0F8FF"  # 薄い青色の背景
        )
        event_setup_frame.pack(fill="x", padx=20, pady=10)

        # 説明ラベル
        tk.Label(
            event_setup_frame,
            text="💡 対象ファイルのみ選択して実行してください（参照元ファイルは不要）",
            font=("メイリオ", 9),
            fg="#0066CC",
            bg="#F0F8FF"
        ).pack(pady=(10, 5), padx=15, anchor="w")

        # ボタンコンテナ
        event_button_container = tk.Frame(event_setup_frame, bg="#F0F8FF")
        event_button_container.pack(pady=(5, 10))

        self.read_event_button = create_hover_button(
            event_button_container,
            text="📥 Excelから行事名を読込",
            command=self._read_event_names_from_excel,
            color="success",
            font=("メイリオ", 9),
            width=26,
            height=2
        )
        self.read_event_button.pack()
        create_tooltip(self.read_event_button, "対象Excelの行事名をアプリ設定に読み込みます（対象ファイルのみで動作）")

        # 補足説明
        tk.Label(
            event_setup_frame,
            text="※ 既存のExcelファイルから行事名リストを取り込みます。設定タブで編集も可能です。",
            font=("メイリオ", 8),
            fg="#666",
            bg="#F0F8FF"
        ).pack(pady=(0, 10), padx=15, anchor="w")

        # 実行ボタン（横並び）
        button_frame = tk.Frame(main_container)
        button_frame.pack(pady=15)

        # メインボタンコンテナ
        main_buttons = tk.Frame(button_frame)
        main_buttons.pack()

        # データ更新ボタン（左側）
        self.run_button = create_hover_button(
            main_buttons,
            text="▶ Excelデータ更新を実行",
            command=self._run_excel_update,
            color="primary",
            font=("メイリオ", 11, "bold"),
            width=28,
            height=2
        )
        self.run_button.pack(side="left", padx=5)

        # ステータスラベル
        self.status_label = tk.Label(main_container, text="", font=("メイリオ", 9), fg="gray")
        self.status_label.pack()

        # ログ表示
        self.create_log_frame(height=8, parent=main_container)
        self.log("✅ 準備完了。「🔍 自動検出」または「📁 ファイルを選択」からExcelファイルを指定してください。", "info")

    def _auto_detect_files(self) -> None:
        """開いているExcelファイルを自動検出"""
        if win32com is None or pythoncom is None:
            messagebox.showerror(
                "モジュールエラー",
                "win32comがインストールされていません。\n\n"
                "以下のコマンドでインストールしてください:\n"
                "pip install pywin32"
            )
            return

        import threading

        # config.jsonから自動検出キーワードを取得
        ref_keywords = self.config.get('excel_auto_detect', 'reference_keywords')
        target_keywords = self.config.get('excel_auto_detect', 'target_keywords')

        # キーワードが取得できない場合はデフォルト値を使用
        if not ref_keywords:
            ref_keywords = ["編集用", "年間行事"]
        if not target_keywords:
            target_keywords = ["様式4", "様式４"]

        # COM操作を別スレッド(MTA)で実行（メインスレッドはSTA固定のため）
        result: Dict[str, list] = {"ref": [], "target": [], "error": []}

        def _detect_in_thread() -> None:
            try:
                pythoncom.CoInitializeEx(pythoncom.COINIT_MULTITHREADED)
                try:
                    excel = win32com.client.Dispatch("Excel.Application")
                    for wb in excel.Workbooks:
                        full_path = wb.FullName
                        filename = wb.Name
                        sheet_names = []
                        try:
                            for ws in wb.Worksheets:
                                sheet_names.append(ws.Name)
                        except Exception:
                            pass
                        search_targets = [filename] + sheet_names
                        if any(kw in t for kw in ref_keywords for t in search_targets):
                            result["ref"].append((filename, full_path))
                        if any(kw in t for kw in target_keywords for t in search_targets):
                            result["target"].append((filename, full_path))
                    del excel
                finally:
                    pythoncom.CoUninitialize()
            except Exception as e:
                result["error"].append(str(e))

        thread = threading.Thread(target=_detect_in_thread, daemon=True)
        thread.start()
        thread.join(timeout=10)

        if result["error"]:
            error_msg = result["error"][0]
            logger.error(f"Excelアプリケーションへの接続エラー: {error_msg}")
            messagebox.showerror(
                "Excel接続エラー",
                "Excelが起動していないか、接続できません。\n\n"
                "以下を確認してください:\n"
                "• Microsoft Excelが起動しているか\n"
                "• 対象ファイルが開いているか"
            )
            return

        ref_candidates = result["ref"]
        target_candidates = result["target"]

        # 自動検出結果をログ出力
        self.log("=== 自動検出開始 ===", "info")

        if ref_candidates:
            selected_ref = ref_candidates[0]
            self.ref_file_path = selected_ref[1]
            self._update_file_label("reference", selected_ref[0], selected_ref[1])
            self.log(f"✅ 参照元を検出: {selected_ref[0]}", "success")
            if len(ref_candidates) > 1:
                self.log(f"   ℹ️ 他の候補: {', '.join([c[0] for c in ref_candidates[1:]])}", "info")
        else:
            keywords_str = "」「".join(ref_keywords)
            self.log(f"❌ 参照元ファイルが見つかりません（ファイル名またはシート名に「{keywords_str}」を含むファイルを開いてください）", "warning")

        if target_candidates:
            selected_target = target_candidates[0]
            self.target_file_path = selected_target[1]
            self._update_file_label("target", selected_target[0], selected_target[1])
            self.log(f"✅ 対象を検出: {selected_target[0]}", "success")
            if len(target_candidates) > 1:
                self.log(f"   ℹ️ 他の候補: {', '.join([c[0] for c in target_candidates[1:]])}", "info")
        else:
            keywords_str = "」「".join(target_keywords)
            self.log(f"❌ 対象ファイルが見つかりません（ファイル名またはシート名に「{keywords_str}」を含むファイルを開いてください）", "warning")

        if ref_candidates and target_candidates:
            self.update_status("✅ 両方のファイルを自動検出しました")
            messagebox.showinfo(
                "検出成功",
                f"参照元: {ref_candidates[0][0]}\n対象: {target_candidates[0][0]}\n\n"
                "ファイルが正しいか確認してください。"
            )
        elif ref_candidates or target_candidates:
            self.update_status("⚠️ 一部のファイルのみ検出されました")
            messagebox.showwarning(
                "一部検出",
                "一部のファイルのみ検出されました。\n手動で残りのファイルを選択してください。"
            )
        else:
            self.update_status("❌ ファイルが検出されませんでした")
            ref_keywords_str = "」「".join(ref_keywords)
            target_keywords_str = "」「".join(target_keywords)
            messagebox.showwarning(
                "未検出",
                "対象ファイルが見つかりませんでした。\n\n"
                f"参照元: 「{ref_keywords_str}」を含むファイル\n"
                f"対象: 「{target_keywords_str}」を含むファイル\n\n"
                "上記ファイルをExcelで開いてから再実行してください。"
            )

    def _select_file(self, file_type: str) -> None:
        """
        ファイル選択ダイアログを表示

        Args:
            file_type: "reference" または "target"
        """
        # 初期ディレクトリを決定
        initial_dir = Path.home() / "Downloads"  # デフォルトはダウンロードフォルダ

        # 現在のファイルパスが存在する場合はそのディレクトリを使用
        current_path = self.ref_file_path if file_type == "reference" else self.target_file_path
        if current_path and os.path.exists(current_path):
            initial_dir = Path(current_path).parent

        file_path = filedialog.askopenfilename(
            title="Excelファイルを選択" + (" (参照元)" if file_type == "reference" else " (対象)"),
            initialdir=str(initial_dir),
            filetypes=[("Excelファイル", "*.xlsx;*.xls"), ("すべて", "*.*")]
        )

        if file_path:
            # PathValidatorで検証
            is_valid, error_msg, validated_path = PathValidator.validate_file_path(
                file_path,
                must_exist=True
            )

            if not is_valid or not validated_path:
                messagebox.showerror("パス検証エラー", error_msg or "ファイルパスが無効です")
                return

            # ファイルパスを保存
            if file_type == "reference":
                self.ref_file_path = str(validated_path)
            else:
                self.target_file_path = str(validated_path)

            # ラベルを更新
            filename = validated_path.name
            self._update_file_label(file_type, filename, str(validated_path))

            self.log(f"✅ {'参照元' if file_type == 'reference' else '対象'}ファイルを選択: {filename}", "success")
            self.update_status(f"✅ ファイルを選択: {filename}")

    def _update_file_label(self, file_type: str, filename: str, full_path: str) -> None:
        """
        ファイルラベルとステータスを更新

        Args:
            file_type: "reference" または "target"
            filename: ファイル名
            full_path: フルパス
        """
        if file_type == "reference":
            self.ref_label.config(text=full_path, fg="#2196F3")
            self.ref_status.config(fg="green")
        else:
            self.target_label.config(text=full_path, fg="#2196F3")
            self.target_status.config(fg="green")

    def _open_selected_file(self, file_path: Optional[str]) -> None:
        """
        選択されたファイルを開く

        Args:
            file_path: 開くファイルのパス
        """
        if not file_path:
            messagebox.showwarning("警告", "ファイルが選択されていません。")
            return

        # PathValidatorでファイルパスを検証
        is_valid, error_msg, validated_path = PathValidator.validate_file_path(
            file_path,
            must_exist=True
        )

        if not is_valid:
            messagebox.showerror("パス検証エラー", error_msg or "ファイルパスが無効です")
            return

        def on_error(error_msg: str) -> None:
            messagebox.showerror("エラー", error_msg)
            self.log(f"❌ Excelファイルを開けませんでした: {file_path}", "error")

        if validated_path and open_file_or_folder(str(validated_path), on_error):
            self.log(f"📂 Excelでファイルを開きました: {Path(file_path).name}", "info")
            self.update_status(f"✅ Excelでファイルを開きました: {Path(file_path).name}")

    def _run_excel_update(self) -> None:
        """Excelデータ更新を実行"""
        # ファイルが選択されているか確認
        if not self.ref_file_path or not self.target_file_path:
            missing = []
            if not self.ref_file_path:
                missing.append("• 参照元ファイル")
            if not self.target_file_path:
                missing.append("• 対象ファイル")

            messagebox.showerror(
                "ファイル未選択",
                "以下のファイルが選択されていません:\n\n" + "\n".join(missing) + "\n\n"
                "「自動検出」または「ファイルを選択」でファイルを指定してください。"
            )
            return

        # ファイルが存在するか確認
        if not os.path.exists(self.ref_file_path):
            messagebox.showerror(
                "ファイルが見つかりません",
                f"参照元ファイルが見つかりません。\n\n"
                f"ファイルパス:\n{self.ref_file_path}\n\n"
                f"ファイルが移動または削除されていないか確認してください。"
            )
            return

        if not os.path.exists(self.target_file_path):
            messagebox.showerror(
                "ファイルが見つかりません",
                f"対象ファイルが見つかりません。\n\n"
                f"ファイルパス:\n{self.target_file_path}\n\n"
                f"ファイルが移動または削除されていないか確認してください。"
            )
            return

        def task():
            try:
                set_button_state(self.run_button, False, self.status_label, "🔄 実行中...")
                self.update_status("🔄 Excelデータ更新を実行中...")
                self.log("=== Excelデータ更新開始 ===", "info")

                # シート名は設定ファイルから取得
                ref_sheet = self.config.get('files', 'excel_reference_sheet')
                target_sheet = self.config.get('files', 'excel_target_sheet')

                self.log(f"参照ファイル: {Path(self.ref_file_path).name}", "info")
                self.log(f"ターゲットファイル: {Path(self.target_file_path).name}", "info")
                self.log(f"参照シート: {ref_sheet}", "info")
                self.log(f"ターゲットシート: {target_sheet}", "info")

                # 進捗コールバック関数を定義
                def progress_callback(message: str) -> None:
                    """進捗状況をGUIに反映"""
                    self.log(message, "info")
                    self.update_status(message)

                # ステップ1: ConfigLoaderから最新の行事名を取得してターゲットExcelに設定
                self.log("📝 ステップ1: 行事名をターゲットExcelに設定中...", "info")
                self.update_status("📝 行事名を設定中...")

                school_events = self.config.get_event_names("school_events")
                student_council_events = self.config.get_event_names("student_council_events")
                other_activities = self.config.get_event_names("other_activities")

                # 行事名を設定するための一時的なExcelTransferインスタンス
                temp_transfer = ExcelTransfer(
                    ref_filename="",  # 行事名設定では参照ファイル不要
                    target_filename=self.target_file_path,
                    ref_sheet=ref_sheet,
                    target_sheet=target_sheet
                )
                counts = temp_transfer.populate_event_names(
                    school_events=school_events,
                    student_council_events=student_council_events,
                    other_activities=other_activities
                )

                total_events = sum(counts.values())
                self.log(
                    f"✅ 行事名を設定しました（学校行事:{counts['school_events']}件、"
                    f"児童会行事:{counts['student_council_events']}件、"
                    f"その他:{counts['other_activities']}件、合計:{total_events}件）",
                    "success"
                )

                # ステップ2: 転記処理を実行
                self.log("🔄 ステップ2: 転記処理を開始します...", "info")
                self.update_status("🔄 転記処理を実行中...")

                # ExcelTransferにフルパスを渡す
                transfer = ExcelTransfer(
                    ref_filename=self.ref_file_path,
                    target_filename=self.target_file_path,
                    ref_sheet=ref_sheet,
                    target_sheet=target_sheet,
                    progress_callback=progress_callback,
                    cancel_check=None
                )
                transfer.execute()

                self.log("✅ === Excelデータ更新完了 ===", "success")
                set_button_state(self.run_button, True, self.status_label, "")
                self.update_status("✅ Excelデータ更新が完了しました")
                thread_safe_call(self.tab, lambda: messagebox.showinfo(
                    "完了",
                    "Excelデータ更新が完了しました。\n\n"
                    "内容を確認して保存してください。"
                ))
            except Exception as e:
                self.log(f"❌ エラー: {e}", "error")
                set_button_state(self.run_button, True, self.status_label, "")
                self.update_status("❌ エラーが発生しました")
                error_msg = str(e)
                thread_safe_call(self.tab, lambda: messagebox.showerror(
                    "実行エラー", f"エラーが発生しました。\n\n詳細:\n{error_msg}"
                ))

        thread = threading.Thread(target=task, daemon=True)
        thread.start()

    def _read_event_names_from_excel(self) -> None:
        """Excelから行事名を読み込み（ワンクリック）"""
        # 1. ターゲットファイルチェック
        if not self.target_file_path:
            self.log("❌ エラー: 対象ファイルを選択してください", "error")
            messagebox.showerror(
                "対象ファイル未選択",
                "対象ファイル（様式4）を選択してください。\n\n"
                "※ 参照元ファイルは不要です。対象ファイルのみ選択してください。"
            )
            return

        # ファイルが存在するか確認
        if not os.path.exists(self.target_file_path):
            self.log(f"❌ エラー: ファイルが見つかりません - {Path(self.target_file_path).name}", "error")
            messagebox.showerror(
                "ファイルが見つかりません",
                f"指定されたファイルが見つかりません。\n\n"
                f"ファイルパス:\n{self.target_file_path}\n\n"
                f"ファイルが移動または削除されていないか確認してください。"
            )
            return

        # 2. 既存設定の有無を確認
        existing_config = self.config.user_config.get("excel_event_names", {})
        has_existing = bool(existing_config)

        # 3. 確認ダイアログ（既存設定の有無で表示を変える）
        if has_existing:
            # 既存設定がある場合：上書き警告を強化
            existing_counts = {
                "学校行事名": len(existing_config.get("school_events", [])),
                "児童会行事名": len(existing_config.get("student_council_events", [])),
                "その他の活動": len(existing_config.get("other_activities", []))
            }
            total_existing = sum(existing_counts.values())

            result = messagebox.askyesno(
                "⚠️ 既存設定の上書き確認",
                f"既に行事名が設定されています（合計: {total_existing}件）\n\n"
                f"現在の設定:\n"
                f"  • 学校行事名: {existing_counts['学校行事名']}件\n"
                f"  • 児童会行事名: {existing_counts['児童会行事名']}件\n"
                f"  • その他の活動: {existing_counts['その他の活動']}件\n\n"
                f"対象Excelから読み込んで上書きしますか？\n\n"
                f"ファイル: {Path(self.target_file_path).name}\n\n"
                f"※ 現在の設定は失われます。",
                icon="warning"
            )
        else:
            # 初回セットアップの場合：シンプルな確認
            result = messagebox.askyesno(
                "📥 初回セットアップ確認",
                f"対象Excelから行事名を読み込みますか？\n\n"
                f"ファイル: {Path(self.target_file_path).name}\n\n"
                f"※ これは初回セットアップです。\n"
                f"※ 読み込み後、設定タブで確認・編集できます。",
                icon="question"
            )

        if not result:
            self.log("ℹ️  行事名読み込みをキャンセルしました", "info")
            return

        # 4. ボタン無効化
        set_button_state(self.read_event_button, False, self.status_label, "読み込み中...")

        if has_existing:
            self.log("🔄 Excelから行事名を読み込んでいます（既存設定を上書き）...", "info")
        else:
            self.log("📥 Excelから行事名を読み込んでいます（初回セットアップ）...", "info")

        def task():
            try:
                # ExcelTransferインスタンス作成（COM管理はExcelTransferに任せる）
                transfer = ExcelTransfer(
                    ref_filename="",  # 行事名読み込みでは参照ファイル不要
                    target_filename=self.target_file_path,
                    ref_sheet=self.config.get("files", "excel_reference_sheet"),
                    target_sheet=self.config.get("files", "excel_target_sheet")
                )

                # Excelから読み込み
                event_data = transfer.read_event_names_from_excel()

                # ConfigLoaderに保存
                for category, event_names in event_data.items():
                    if event_names:  # 空でない場合のみ保存
                        self.config.save_event_names(category, event_names)

                # 件数を計算
                counts = {k: len(v) for k, v in event_data.items()}
                total = sum(counts.values())

                # 成功メッセージ（既存設定の有無で変更）
                if has_existing:
                    self.log(
                        f"✅ 行事名を上書きしました（合計: {total}件）",
                        "success"
                    )
                    self.update_status(f"✅ 行事名を{total}件上書きしました")
                    dialog_title = "✅ 上書き完了"
                    dialog_message = (
                        f"既存の設定を上書きしました。\n\n"
                        f"学校行事名: {counts['school_events']}件\n"
                        f"児童会行事名: {counts['student_council_events']}件\n"
                        f"その他の活動: {counts['other_activities']}件\n"
                        f"合計: {total}件\n\n"
                        f"設定タブで確認・編集できます。"
                    )
                else:
                    self.log(
                        f"✅ 行事名を読み込みました（初回セットアップ完了、合計: {total}件）",
                        "success"
                    )
                    self.update_status(f"✅ 初回セットアップ完了（{total}件）")
                    dialog_title = "📥 初回セットアップ完了"
                    dialog_message = (
                        f"Excelから行事名を読み込みました。\n\n"
                        f"学校行事名: {counts['school_events']}件\n"
                        f"児童会行事名: {counts['student_council_events']}件\n"
                        f"その他の活動: {counts['other_activities']}件\n"
                        f"合計: {total}件\n\n"
                        f"設定タブで確認・編集できます。"
                    )

                # 詳細ダイアログ（メインスレッドで表示）
                thread_safe_call(self.tab, lambda: messagebox.showinfo(dialog_title, dialog_message))

                # 設定タブの行事名リストをリロード（メインスレッドで実行）
                if self.settings_tab:
                    self.tab.after(0, self.settings_tab.reload_event_names)
                    logger.info("設定タブに行事名の更新を通知しました")

                # ボタン再有効化
                set_button_state(self.read_event_button, True, self.status_label, "")

            except Exception as e:
                logger.error(f"行事名読み込みエラー: {e}", exc_info=True)
                self.log(f"❌ エラー: {e}", "error")
                set_button_state(self.read_event_button, True, self.status_label, "")
                self.update_status("❌ エラーが発生しました")
                error_msg = str(e)
                thread_safe_call(self.tab, lambda: messagebox.showerror(
                    "読み込みエラー",
                    f"行事名の読み込みに失敗しました。\n\n詳細:\n{error_msg}"
                ))

        thread = threading.Thread(target=task, daemon=True)
        thread.start()
