"""
Excel処理タブ

Excel自動更新機能のUIを提供
"""
import logging
import os
import tkinter as tk
from tkinter import ttk, messagebox
import threading
from typing import Tuple, TYPE_CHECKING

try:
    import win32com.client
    import pythoncom
except ImportError:
    win32com = None
    pythoncom = None

from gui.tabs.base_tab import BaseTab
from gui.utils import set_button_state, create_hover_button, open_file_or_folder, create_tooltip
from path_validator import PathValidator
from transfer_factory import HybridTransferFactory

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
        self._create_ui()
        self.add_to_notebook("📊 Excel処理")

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
            "📝 実行手順（必ずこの順番で操作してください）\n\n"
            "1️⃣ 下記の2つのExcelファイルを「Excelアプリ」で開く\n"
            "   ※ ファイル名をダブルクリックすると開きます\n"
            "   ※ 「📁 開く」ボタンでも開けます\n\n"
            "2️⃣ ●マークが緑色になったことを確認\n"
            "   ※ 緑色 = ファイルが開いています\n"
            "   ※ F5キーで状態を更新できます\n\n"
            "3️⃣ 「Excelデータ更新を実行」ボタンをクリック\n\n"
            "4️⃣ 処理完了後、内容を確認してExcelで保存"
        )
        tk.Label(info_frame, text=steps_text, justify="left", font=("メイリオ", 9), fg="#333").pack(pady=(5, 15), padx=15, anchor="w")

        # ファイル選択フレーム
        file_frame = tk.LabelFrame(main_container, text="📂 対象ファイル", font=("メイリオ", 11, "bold"))
        file_frame.pack(fill="x", padx=20, pady=10)

        # 参照元ファイル
        ref_frame = tk.Frame(file_frame)
        ref_frame.pack(fill="x", padx=15, pady=8)

        tk.Label(ref_frame, text="参照元:", width=12, anchor="w", font=("メイリオ", 9, "bold")).pack(side="left")
        self.ref_label = tk.Label(
            ref_frame,
            text=self.config.get('files', 'excel_reference'),
            font=("メイリオ", 9),
            fg="#2196F3",
            anchor="w",
            cursor="hand2"
        )
        self.ref_label.pack(side="left", fill="x", expand=True, padx=10)
        self.ref_label.bind("<Button-1>", lambda e: self._open_excel_file(self.config.get('files', 'excel_reference')))
        create_tooltip(self.ref_label, "クリックでExcelファイルを開きます")

        ref_open_btn = tk.Button(
            ref_frame,
            text="📁 開く",
            command=lambda: self._open_excel_file(self.config.get('files', 'excel_reference')),
            width=8,
            font=("メイリオ", 9)
        )
        ref_open_btn.pack(side="right", padx=2)
        create_tooltip(ref_open_btn, "参照元ファイルをExcelで開きます")

        self.ref_status = tk.Label(ref_frame, text="●", fg="gray", font=("メイリオ", 12))
        self.ref_status.pack(side="right")
        create_tooltip(self.ref_status, "緑色 = ファイルが開いています")

        # 対象ファイル
        target_frame = tk.Frame(file_frame)
        target_frame.pack(fill="x", padx=15, pady=8)

        tk.Label(target_frame, text="対象:", width=12, anchor="w", font=("メイリオ", 9, "bold")).pack(side="left")
        self.target_label = tk.Label(
            target_frame,
            text=self.config.get('files', 'excel_target'),
            font=("メイリオ", 9),
            fg="#2196F3",
            anchor="w",
            cursor="hand2"
        )
        self.target_label.pack(side="left", fill="x", expand=True, padx=10)
        self.target_label.bind("<Button-1>", lambda e: self._open_excel_file(self.config.get('files', 'excel_target')))
        create_tooltip(self.target_label, "クリックでExcelファイルを開きます")

        target_open_btn = tk.Button(
            target_frame,
            text="📁 開く",
            command=lambda: self._open_excel_file(self.config.get('files', 'excel_target')),
            width=8,
            font=("メイリオ", 9)
        )
        target_open_btn.pack(side="right", padx=2)
        create_tooltip(target_open_btn, "対象ファイルをExcelで開きます")

        self.target_status = tk.Label(target_frame, text="●", fg="gray", font=("メイリオ", 12))
        self.target_status.pack(side="right")
        create_tooltip(self.target_status, "緑色 = ファイルが開いています")

        # ファイル状態確認ボタン
        check_frame = tk.Frame(file_frame)
        check_frame.pack(pady=10)
        check_btn = tk.Button(
            check_frame,
            text="🔄 ファイル状態を確認 (F5)",
            command=self.check_files_status,
            font=("メイリオ", 9),
            width=25
        )
        check_btn.pack()
        create_tooltip(check_btn, "●マークの色を更新します（F5キーでも可）")

        # 実行ボタン（現在は無効化）
        button_frame = tk.Frame(main_container)
        button_frame.pack(pady=15)

        self.run_button = create_hover_button(
            button_frame,
            text="▶ Excelデータ更新を実行",
            command=self._run_excel_update,
            color="primary",
            font=("メイリオ", 11, "bold"),
            width=32,
            height=2
        )
        self.run_button.pack()

        # 使用方法の説明
        tk.Label(
            button_frame,
            text="※ 両方のExcelファイルを開いてから実行してください",
            font=("メイリオ", 9),
            fg="#0288D1"  # Material Light Blue 700
        ).pack(pady=(5, 0))

        # ステータスラベル
        self.status_label = tk.Label(main_container, text="", font=("メイリオ", 9), fg="gray")
        self.status_label.pack()

        # ログ表示
        self.create_log_frame(height=8, parent=main_container)
        self.log("準備完了。上記の2つのExcelファイルを開いてから実行してください。", "info")

    def _open_excel_file(self, filename: str) -> None:
        """
        Excelファイルを開く

        Args:
            filename: 開くファイル名（相対パスまたは絶対パス）
        """
        # 絶対パス（C:\, \\, /で始まる）の場合はそのまま使用
        if os.path.isabs(filename) or filename.startswith('\\\\') or filename.startswith('//'):
            file_path = filename
        else:
            # 相対パスの場合は従来通りパス構築
            base_path = self.config.get('base_paths', 'google_drive')
            year = self.config.year
            year_short = self.config.year_short
            education_base = self.config.get('directories', 'education_plan_base')

            # {year_short}プレースホルダーを実際の値に置き換える
            education_base = education_base.replace('{year_short}', year_short)

            file_path = os.path.join(base_path, year, education_base, filename)

        # PathValidatorでファイルパスを検証
        is_valid, error_msg, validated_path = PathValidator.validate_file_path(
            file_path,
            must_exist=False  # ファイルが存在しなくてもエラーメッセージで通知
        )
        if not is_valid:
            messagebox.showerror("パス検証エラー", error_msg)
            self.log(f"ファイルパスが無効です: {error_msg}", "error")
            return

        def on_error(error_msg: str) -> None:
            messagebox.showerror("エラー", error_msg)
            self.log(f"Excelファイルを開けませんでした: {filename}", "error")

        if open_file_or_folder(str(validated_path), on_error):
            self.log(f"Excelでファイルを開きました: {filename}", "info")
            self.update_status(f"Excelでファイルを開きました: {filename}")

    def _check_excel_files_open(self) -> Tuple[bool, bool]:
        """
        Excelファイルが開かれているか確認

        Returns:
            Tuple[bool, bool]: (参照元が開いているか, 対象が開いているか)
        """
        # win32comが未インストールの場合は即座に返す
        if win32com is None or pythoncom is None:
            logger.error("win32comがインストールされていません")
            return False, False

        excel = None
        com_initialized = False
        try:
            # COM初期化（失敗時のクリーンアップのためフラグ管理）
            try:
                pythoncom.CoInitialize()
                com_initialized = True
            except Exception as e:
                logger.error(f"COM初期化エラー: {e}")
                return False, False

            # Excelアプリケーションに接続
            try:
                excel = win32com.client.Dispatch("Excel.Application")
            except Exception as e:
                logger.error(f"Excelアプリケーションへの接続エラー: {e}")
                return False, False

            ref_filename = self.config.get('files', 'excel_reference')
            target_filename = self.config.get('files', 'excel_target')

            # フルパスの場合はファイル名のみを抽出
            ref_basename = os.path.basename(ref_filename) if ref_filename else ""
            target_basename = os.path.basename(target_filename) if target_filename else ""

            ref_open = False
            target_open = False

            for wb in excel.Workbooks:
                if ref_basename and ref_basename in wb.Name:
                    ref_open = True
                if target_basename and target_basename in wb.Name:
                    target_open = True

            return ref_open, target_open

        except Exception as e:
            logger.warning(f"Excel状態チェックエラー: {e}")
            return False, False

        finally:
            # COMオブジェクトを完全に解放（例外時も必ず実行）
            if excel is not None:
                try:
                    del excel
                except Exception as e:
                    logger.warning(f"COMオブジェクト削除エラー: {e}")
                excel = None

            # COM終了処理（初期化成功時のみ実行）
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception as cleanup_error:
                    logger.warning(f"COM終了処理エラー: {cleanup_error}")

    def check_files_status(self) -> None:
        """Excelファイルの状態を確認"""
        try:
            ref_open, target_open = self._check_excel_files_open()

            self.ref_status.config(text="●", fg="green" if ref_open else "gray")
            self.target_status.config(text="●", fg="green" if target_open else "gray")

            if ref_open and target_open:
                self.log("✅ 両方のファイルが開かれています", "success")
                self.update_status("✅ 両方のExcelファイルが開かれています")
            elif ref_open:
                self.log("⚠️ 参照元ファイルのみ開かれています。対象ファイルも開いてください。", "warning")
                self.update_status("⚠️ 参照元ファイルのみ開いています")
            elif target_open:
                self.log("⚠️ 対象ファイルのみ開かれています。参照元ファイルも開いてください。", "warning")
                self.update_status("⚠️ 対象ファイルのみ開いています")
            else:
                self.log("❌ どちらのファイルも開かれていません", "error")
                self.update_status("❌ Excelファイルが開かれていません")

        except Exception as e:
            self.log(f"ファイル状態の確認に失敗: {e}", "error")
            messagebox.showerror("確認エラー", f"Excelファイルの状態を確認できませんでした。\n\n詳細: {e}")

    def update_labels(self) -> None:
        """ラベルを更新"""
        self.ref_label.config(text=self.config.get('files', 'excel_reference'))
        self.target_label.config(text=self.config.get('files', 'excel_target'))

    def _run_excel_update(self) -> None:
        """Excelデータ更新を実行"""
        # 参照モードを取得
        reference_mode = self.config.get('files', 'reference_mode', default='excel')

        # 実行前にファイル状態を確認
        ref_open, target_open = self._check_excel_files_open()

        # モード別のファイルチェック
        if reference_mode == 'excel':
            # Excelモード: 両方のファイルが開かれている必要がある
            if not (ref_open and target_open):
                ref_filename = self.config.get('files', 'excel_reference')
                target_filename = self.config.get('files', 'excel_target')

                missing = []
                if not ref_open:
                    missing.append(f"• {ref_filename}")
                if not target_open:
                    missing.append(f"• {target_filename}")

                result = messagebox.askokcancel(
                    "ファイル未オープン",
                    "以下のファイルが開かれていません:\n\n" + "\n".join(missing) + "\n\n続行しますか？"
                )
                if not result:
                    return
        elif reference_mode == 'google_sheets':
            # Google Sheetsモード: ターゲットファイルのみ開かれている必要がある
            if not target_open:
                target_filename = self.config.get('files', 'excel_target')

                result = messagebox.askokcancel(
                    "ファイル未オープン",
                    f"ターゲットファイルが開かれていません:\n\n• {target_filename}\n\n続行しますか？"
                )
                if not result:
                    return

        def task():
            try:
                set_button_state(self.run_button, False, self.status_label, "🔄 実行中...")
                self.update_status("Excelデータ更新を実行中...")
                self.log("=== Excelデータ更新開始 ===", "info")

                # 参照モードを取得
                reference_mode = self.config.get('files', 'reference_mode', default='excel')
                self.log(f"参照モード: {reference_mode}", "info")

                # ターゲットファイル情報をログ出力
                target_filename = self.config.get('files', 'excel_target')
                self.log(f"ターゲットファイル: {target_filename}", "info")

                # 進捗コールバック関数を定義
                def progress_callback(message: str) -> None:
                    """進捗状況をGUIに反映"""
                    self.log(message, "info")
                    self.update_status(message)

                # 転送インスタンスを生成（ファクトリーパターン）
                transfer = HybridTransferFactory.create_transfer(
                    config=self.config,
                    progress_callback=progress_callback,
                    cancel_check=None
                )
                transfer.execute()

                self.log("=== Excelデータ更新完了 ===", "success")
                set_button_state(self.run_button, True, self.status_label, "")
                self.update_status("Excelデータ更新が完了しました")
                messagebox.showinfo(
                    "完了",
                    "Excelデータ更新が完了しました。\n\n"
                    "内容を確認して保存してください。"
                )
            except Exception as e:
                self.log(f"エラー: {e}", "error")
                set_button_state(self.run_button, True, self.status_label, "")
                self.update_status("エラーが発生しました")
                messagebox.showerror("実行エラー", f"エラーが発生しました。\n\n詳細:\n{e}")

        thread = threading.Thread(target=task, daemon=True)
        thread.start()
