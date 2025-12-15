"""
Excel処理タブ

Excel自動更新機能のUIを提供
"""
import logging
import os
import tkinter as tk
from tkinter import ttk, messagebox
import threading
from typing import Any

from gui.tabs.base_tab import BaseTab
from gui.utils import set_button_state, create_hover_button

# ロガーの設定
logger = logging.getLogger(__name__)


class ExcelTab(BaseTab):
    """Excel処理タブ"""

    def __init__(self, notebook: ttk.Notebook, config: Any, status_bar: tk.Label) -> None:
        super().__init__(notebook, config, status_bar)
        self._create_ui()
        self.add_to_notebook("📊 Excel処理")

    def _create_ui(self) -> None:
        """UIを構築"""
        # 説明フレーム
        info_frame = tk.LabelFrame(self.tab, text="📋 Excel自動更新機能", font=("メイリオ", 11, "bold"))
        info_frame.pack(fill="x", padx=20, pady=15)

        info_text = "年間行事計画（編集用）から様式4へ自動的にデータを転記します。"
        tk.Label(info_frame, text=info_text, justify="left", font=("メイリオ", 10)).pack(pady=(15, 5), padx=15)

        steps_text = "実行手順:\n1️⃣ 下記の2つのExcelファイルをExcelで開く\n2️⃣ 「Excelデータ更新を実行」ボタンをクリック\n3️⃣ 処理完了後、内容を確認して保存"
        tk.Label(info_frame, text=steps_text, justify="left", font=("メイリオ", 9), fg="#555").pack(pady=(5, 15), padx=15, anchor="w")

        # ファイル選択フレーム
        file_frame = tk.LabelFrame(self.tab, text="📂 対象ファイル", font=("メイリオ", 11, "bold"))
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

        tk.Button(
            ref_frame,
            text="📁 開く",
            command=lambda: self._open_excel_file(self.config.get('files', 'excel_reference')),
            width=8,
            font=("メイリオ", 9)
        ).pack(side="right", padx=2)

        self.ref_status = tk.Label(ref_frame, text="●", fg="gray", font=("メイリオ", 12))
        self.ref_status.pack(side="right")

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

        tk.Button(
            target_frame,
            text="📁 開く",
            command=lambda: self._open_excel_file(self.config.get('files', 'excel_target')),
            width=8,
            font=("メイリオ", 9)
        ).pack(side="right", padx=2)

        self.target_status = tk.Label(target_frame, text="●", fg="gray", font=("メイリオ", 12))
        self.target_status.pack(side="right")

        # ファイル状態確認ボタン
        check_frame = tk.Frame(file_frame)
        check_frame.pack(pady=10)
        tk.Button(
            check_frame,
            text="🔄 ファイル状態を確認 (F5)",
            command=self.check_files_status,
            font=("メイリオ", 9),
            width=25
        ).pack()

        # 実行ボタン
        button_frame = tk.Frame(self.tab)
        button_frame.pack(pady=15)

        self.run_button = create_hover_button(
            button_frame,
            text="▶ Excelデータ更新を実行",
            command=self._run_excel_update,
            color="secondary",
            font=("メイリオ", 11, "bold"),
            width=32,
            height=2
        )
        self.run_button.pack()

        # ステータスラベル
        self.status_label = tk.Label(self.tab, text="", font=("メイリオ", 9), fg="gray")
        self.status_label.pack()

        # ログ表示
        self.create_log_frame(height=8)
        self.log("準備完了。上記の2つのExcelファイルを開いてから実行してください。", "info")

    def _open_excel_file(self, filename: str) -> None:
        """Excelファイルを開く"""
        try:
            base_path = self.config.get('base_paths', 'google_drive')
            year = self.config.year
            education_base = self.config.get('directories', 'education_plan_base')
            file_path = os.path.join(base_path, year, education_base, filename)

            if not os.path.exists(file_path):
                messagebox.showerror(
                    "❌ ファイルが見つかりません",
                    f"以下のファイルが見つかりません:\n\n{filename}\n\nパス:\n{file_path}"
                )
                return

            os.startfile(file_path)
            self.log(f"Excelでファイルを開きました: {filename}", "info")
            self.update_status(f"Excelでファイルを開きました: {filename}")

        except Exception as e:
            messagebox.showerror("❌ ファイルオープンエラー", f"ファイルを開けませんでした。\n\n詳細: {e}")

    def check_files_status(self) -> None:
        """Excelファイルの状態を確認"""
        try:
            import win32com.client
            excel = win32com.client.Dispatch("Excel.Application")

            ref_filename = self.config.get('files', 'excel_reference')
            target_filename = self.config.get('files', 'excel_target')

            ref_open = False
            target_open = False

            for wb in excel.Workbooks:
                if ref_filename in wb.Name:
                    ref_open = True
                if target_filename in wb.Name:
                    target_open = True

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
        # 実行前にファイル状態を確認
        try:
            import win32com.client
            excel = win32com.client.Dispatch("Excel.Application")

            ref_filename = self.config.get('files', 'excel_reference')
            target_filename = self.config.get('files', 'excel_target')

            ref_open = False
            target_open = False

            for wb in excel.Workbooks:
                if ref_filename in wb.Name:
                    ref_open = True
                if target_filename in wb.Name:
                    target_open = True

            if not (ref_open and target_open):
                missing = []
                if not ref_open:
                    missing.append(f"• {ref_filename}")
                if not target_open:
                    missing.append(f"• {target_filename}")

                result = messagebox.askokcancel(
                    "⚠️ ファイル未オープン",
                    f"以下のファイルが開かれていません:\n\n" + "\n".join(missing) + "\n\n続行しますか？"
                )
                if not result:
                    return

        except Exception as e:
            # Excelアプリケーションへの接続失敗等（Excelが起動していない場合など）
            logger.debug(f"Excelファイル状態の事前確認をスキップ: {e}")

        def task():
            try:
                set_button_state(self.run_button, False, self.status_label, "🔄 実行中...")
                self.update_status("Excelデータ更新を実行中...")
                self.log("=== Excelデータ更新開始 ===", "info")

                import update_excel_files
                update_excel_files.main()

                self.log("=== Excelデータ更新完了 ===", "success")
                set_button_state(self.run_button, True, self.status_label, "✅ 完了")
                self.update_status("Excelデータ更新が完了しました")
                messagebox.showinfo("✅ 完了", "Excelデータ更新が完了しました！\n\n内容を確認して保存してください。")
            except Exception as e:
                self.log(f"エラー: {e}", "error")
                set_button_state(self.run_button, True, self.status_label, "❌ エラー")
                self.update_status("Excelデータ更新でエラーが発生しました")
                messagebox.showerror("❌ 実行エラー", f"エラーが発生しました。\n\n詳細:\n{e}")

        thread = threading.Thread(target=task, daemon=True)
        thread.start()
