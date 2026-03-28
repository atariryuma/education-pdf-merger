"""
初回セットアップウィザード

ベストプラクティス:
- ステップバイステップのガイド
- プログレスバーによる進捗表示
- 入力検証とエラーメッセージ
- 自動検出機能の活用

参考:
- https://www.kryshiggins.com/the-design-of-setup-wizards/
- https://blog.logrocket.com/ux-design/creating-setup-wizard-when-you-shouldnt/
"""
import logging
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Optional, Callable
from pathlib import Path

from infrastructure.config_loader import ConfigLoader
from infrastructure.ghostscript import GhostscriptDetector
from gui.styles import COLORS, FONTS
from gui.utils import center_window
from infrastructure.path_validator import PathValidator
from infrastructure.year_utils import calculate_next_fiscal_year, calculate_year_short

logger = logging.getLogger(__name__)


class SetupWizard:
    """初回セットアップウィザード

    3ステップのウィザードで基本設定を完了:
    1. ようこそ画面（機能紹介）
    2. 基本設定（年度、作業フォルダ）
    3. 完了画面（設定サマリー、自動検出結果）

    自動設定項目（バックグラウンド）:
    - Ghostscript（自動検出）
    - 一時フォルダ（デフォルト: temp_pdfs）
    - Excel設定（後から設定タブで設定可能）
    """

    def __init__(
        self,
        parent: tk.Tk,
        config: ConfigLoader,
        on_complete: Optional[Callable[[], None]] = None,
    ) -> None:
        """
        初期化

        Args:
            parent: 親ウィンドウ
            config: 設定オブジェクト
            on_complete: 完了時のコールバック
        """
        self.parent = parent
        self.config = config
        self.on_complete = on_complete

        # ウィンドウ作成
        self.window = tk.Toplevel(parent)
        self.window.title("初回セットアップ")
        # 初期サイズは設定せず、内容に合わせて自動調整
        self.window.minsize(700, 600)
        self.window.resizable(True, True)

        # モーダルダイアログとして設定
        self.window.transient(parent)
        self.window.grab_set()

        # 設定値の保持
        default_year, default_year_short = calculate_next_fiscal_year()
        self.year_var = tk.StringVar(value=default_year)
        # year_shortは自動計算（ユーザー入力不要）
        self.year_short_var = tk.StringVar(value=default_year_short)
        self.gdrive_var = tk.StringVar(value="")
        # 自動設定項目（ユーザー入力不要）
        self.local_temp_var = tk.StringVar(value="")  # デフォルト: temp_pdfs
        # v3.5.0: Excelファイルパスは削除（セッション内管理に変更）
        self.gs_var = tk.StringVar(value="")  # 自動検出

        # 現在のステップ
        self.current_step = 0
        self.total_steps = 3

        # UIコンポーネント
        self.content_frame = None
        self.progress_var = tk.IntVar(value=0)

        # 年度変更時に自動でyear_shortを更新
        self.year_var.trace_add("write", self._on_year_changed)

        # UI構築
        self._create_ui()

        # 最初のステップを表示
        self._show_step(0)

        # ウィンドウサイズを内容に合わせて調整してから中央配置
        self.window.update_idletasks()
        center_window(self.window)

        # Ghostscript自動検出（バックグラウンド）
        self.window.after(100, self._detect_ghostscript_async)

    def _on_year_changed(self, *args) -> None:
        """年度が変更されたときに和暦を自動更新"""
        year = self.year_var.get()
        if year.isdigit() and len(year) == 4:
            year_short = calculate_year_short(year)
            self.year_short_var.set(year_short)

    def _create_ui(self) -> None:
        """UI構築"""
        # ヘッダー
        header_frame = tk.Frame(self.window, bg=COLORS["primary"])
        header_frame.pack(fill=tk.X)

        title_label = tk.Label(
            header_frame,
            text="教育計画PDFマージシステム - 初回セットアップ",
            font=FONTS["dialog_heading"],
            bg=COLORS["primary"],
            fg="white",
        )
        title_label.pack(pady=15)

        # プログレスバー
        progress_frame = tk.Frame(self.window, bg="white")
        progress_frame.pack(fill=tk.X)

        self.progress_label = tk.Label(
            progress_frame,
            text=f"ステップ 1 / {self.total_steps}",
            font=FONTS["default"],
            bg="white",
        )
        self.progress_label.pack(pady=2)

        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=self.total_steps,
            mode="determinate",
        )
        self.progress_bar.pack(fill=tk.X, padx=20, pady=5)

        # コンテンツエリア
        self.content_frame = tk.Frame(self.window, bg="white")
        self.content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # ボタンフレーム
        button_frame = tk.Frame(self.window, bg="white")
        button_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=20, pady=10)

        self.back_button = ttk.Button(
            button_frame, text="← 戻る", command=self._go_back, state=tk.DISABLED
        )
        self.back_button.pack(side=tk.LEFT)

        self.next_button = ttk.Button(
            button_frame, text="次へ →", command=self._go_next
        )
        self.next_button.pack(side=tk.RIGHT)

        self.cancel_button = ttk.Button(
            button_frame, text="キャンセル", command=self._cancel
        )
        self.cancel_button.pack(side=tk.RIGHT, padx=10)

    def _show_step(self, step: int) -> None:
        """指定されたステップを表示

        Args:
            step: ステップ番号（0-2: ようこそ、基本設定、完了）
        """
        # コンテンツをクリア
        for widget in self.content_frame.winfo_children():
            widget.destroy()

        self.current_step = step
        self.progress_var.set(step + 1)
        self.progress_label.config(text=f"ステップ {step + 1} / {self.total_steps}")

        # ステップごとの表示
        if step == 0:
            self._show_welcome()
        elif step == 1:
            self._show_basic_settings()  # 年度 + 作業フォルダを統合
        elif step == 2:
            self._show_complete()

        # ボタンの状態更新
        self._update_buttons()

    def _show_welcome(self) -> None:
        """ステップ1: ようこそ画面"""
        # タイトル
        title = tk.Label(
            self.content_frame,
            text="ようこそ！",
            font=FONTS["dialog_title"],
            bg="white",
        )
        title.pack(pady=20)

        # 説明
        desc = tk.Label(
            self.content_frame,
            text=(
                "教育計画PDFマージシステムへようこそ！\n\n"
                "このウィザードでは、アプリケーションを使い始めるために\n"
                "必要な基本設定を行います。\n\n"
                "設定は後から変更することもできます。"
            ),
            font=FONTS["default"],
            bg="white",
            justify=tk.LEFT,
        )
        desc.pack(pady=20)

        # 機能紹介
        features_frame = tk.LabelFrame(
            self.content_frame,
            text="主な機能",
            font=FONTS["default_bold"],
            bg="white",
            relief=tk.GROOVE,
            borderwidth=2,
        )
        features_frame.pack(fill=tk.BOTH, expand=True, pady=20)

        features = [
            "📄 Word・Excel・PowerPointのPDF変換",
            "🖼️ 画像ファイルのPDF変換",
            "📝 一太郎文書のPDF変換",
            "📁 フォルダ構造の自動認識",
            "📊 Excel自動転記機能",
            "🗜️ PDF圧縮機能（Ghostscript）",
        ]

        for feature in features:
            label = tk.Label(
                features_frame,
                text=feature,
                font=FONTS["default"],
                bg="white",
                anchor=tk.W,
            )
            label.pack(fill=tk.X, padx=20, pady=5)

        # 注意事項
        note = tk.Label(
            self.content_frame,
            text="※ Microsoft Officeがインストールされている必要があります",
            font=FONTS["small"],
            bg="white",
            fg="gray",
        )
        note.pack(pady=10)

    def _show_basic_settings(self) -> None:
        """ステップ2: 基本設定（年度 + 作業フォルダ）"""
        # タイトル
        title = tk.Label(
            self.content_frame, text="基本設定", font=FONTS["dialog_title"], bg="white"
        )
        title.pack(pady=15)

        # === 年度設定セクション ===
        year_section = tk.LabelFrame(
            self.content_frame,
            text="📅 年度設定",
            font=FONTS["subheading"],
            bg="white",
            fg=COLORS["primary"],
            relief=tk.GROOVE,
            borderwidth=2,
        )
        year_section.pack(fill=tk.X, padx=20, pady=10)

        # 年度入力（西暦のみ、和暦は自動計算）
        year_frame = tk.Frame(year_section, bg="white")
        year_frame.pack(fill=tk.X, padx=15, pady=8)

        year_label = tk.Label(
            year_frame,
            text="年度（西暦）:",
            font=FONTS["default"],
            bg="white",
            width=12,
            anchor=tk.W,
        )
        year_label.pack(side=tk.LEFT, padx=5)

        year_entry = ttk.Entry(
            year_frame, textvariable=self.year_var, font=FONTS["default"], width=15
        )
        year_entry.pack(side=tk.LEFT, padx=5)

        # 和暦は自動計算される旨を表示（読み取り専用・動的更新）
        arrow_label = tk.Label(year_frame, text="→", font=FONTS["default"], bg="white")
        arrow_label.pack(side=tk.LEFT, padx=5)

        year_short_display_label = tk.Label(
            year_frame,
            textvariable=self.year_short_var,
            font=FONTS["default_bold"],
            bg="white",
            fg=COLORS["primary"],
        )
        year_short_display_label.pack(side=tk.LEFT, padx=5)

        hint_label = tk.Label(
            year_section,
            text="💡 和暦（R8など）は自動計算されます",
            font=FONTS["small"],
            bg="white",
            fg="gray",
        )
        hint_label.pack(padx=15, pady=(0, 8))

        # === 作業フォルダセクション ===
        folder_section = tk.LabelFrame(
            self.content_frame,
            text="📁 作業フォルダ設定",
            font=FONTS["subheading"],
            bg="white",
            fg=COLORS["primary"],
            relief=tk.GROOVE,
            borderwidth=2,
        )
        folder_section.pack(fill=tk.X, padx=20, pady=10)

        desc_label = tk.Label(
            folder_section,
            text="教育計画ファイルが保存されているフォルダを指定してください。",
            font=FONTS["small"],
            bg="white",
            fg="gray",
        )
        desc_label.pack(padx=15, pady=5)

        # フォルダ選択
        folder_frame = tk.Frame(folder_section, bg="white")
        folder_frame.pack(fill=tk.X, padx=15, pady=10)

        folder_label = tk.Label(
            folder_frame,
            text="フォルダ:",
            font=FONTS["default"],
            bg="white",
            width=10,
            anchor=tk.W,
        )
        folder_label.pack(side=tk.LEFT, padx=5)

        folder_entry = ttk.Entry(
            folder_frame, textvariable=self.gdrive_var, font=FONTS["default"], width=35
        )
        folder_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        browse_button = ttk.Button(
            folder_frame, text="参照...", command=self._browse_folder
        )
        browse_button.pack(side=tk.LEFT, padx=5)

        # 状態表示
        self.folder_status_label = tk.Label(
            folder_section, text="", font=FONTS["small"], bg="white", fg="gray"
        )
        self.folder_status_label.pack(padx=15, pady=5)

    def _show_complete(self) -> None:
        """ステップ3: 完了画面"""
        # タイトル
        title = tk.Label(
            self.content_frame,
            text="セットアップ完了！",
            font=FONTS["dialog_title"],
            bg="white",
        )
        title.pack(pady=30)

        # 成功メッセージ
        message = tk.Label(
            self.content_frame,
            text="基本設定が完了しました。\nアプリケーションを使い始めることができます。",
            font=FONTS["default"],
            bg="white",
            justify=tk.CENTER,
        )
        message.pack(pady=20)

        # 設定サマリー
        summary_frame = tk.LabelFrame(
            self.content_frame,
            text="設定内容",
            font=FONTS["default_bold"],
            bg="white",
            relief=tk.GROOVE,
            borderwidth=2,
        )
        summary_frame.pack(fill=tk.BOTH, expand=True, pady=20, padx=20)

        # 年度
        year_label = tk.Label(
            summary_frame,
            text=f"年度: {self.year_var.get()}",
            font=FONTS["default"],
            bg="white",
            anchor=tk.W,
        )
        year_label.pack(fill=tk.X, padx=20, pady=5)

        # 作業フォルダ
        folder_text = self.gdrive_var.get() if self.gdrive_var.get() else "（未設定）"
        folder_label = tk.Label(
            summary_frame,
            text=f"作業フォルダ: {folder_text}",
            font=FONTS["default"],
            bg="white",
            anchor=tk.W,
        )
        folder_label.pack(fill=tk.X, padx=20, pady=5)

        # 自動設定項目
        auto_section = tk.Label(
            self.content_frame,
            text="✨ 自動設定済み",
            font=FONTS["heading"],
            bg="white",
            fg=COLORS["success"],
        )
        auto_section.pack(pady=(20, 10))

        # Ghostscript
        gs_text = (
            f"検出: {self.gs_var.get()}"
            if self.gs_var.get()
            else "未検出（後で設定可能）"
        )
        gs_label = tk.Label(
            self.content_frame,
            text=f"• PDF圧縮機能 (Ghostscript): {gs_text}",
            font=FONTS["small"],
            bg="white",
            anchor=tk.W,
        )
        gs_label.pack(fill=tk.X, padx=40, pady=2)

        # 一時フォルダ
        temp_label = tk.Label(
            self.content_frame,
            text="• 一時フォルダ: デフォルト (temp_pdfs)",
            font=FONTS["small"],
            bg="white",
            anchor=tk.W,
        )
        temp_label.pack(fill=tk.X, padx=40, pady=2)

        # Excel設定
        excel_label = tk.Label(
            self.content_frame,
            text="• Excel自動転記: 設定タブで後から設定可能",
            font=FONTS["small"],
            bg="white",
            anchor=tk.W,
        )
        excel_label.pack(fill=tk.X, padx=40, pady=2)

        # 次のステップ
        next_steps = tk.Label(
            self.content_frame,
            text="設定は「⚙️ 設定」タブからいつでも変更できます。",
            font=FONTS["small"],
            bg="white",
            fg="gray",
        )
        next_steps.pack(pady=20)

    def _update_buttons(self) -> None:
        """ボタンの状態を更新"""
        # 戻るボタン
        if self.current_step == 0:
            self.back_button.config(state=tk.DISABLED)
        else:
            self.back_button.config(state=tk.NORMAL)

        # 次へ/完了ボタン
        if self.current_step == self.total_steps - 1:
            self.next_button.config(text="完了して開始 →")
        else:
            self.next_button.config(text="次へ →")

        # キャンセルボタン（最終ステップでは非表示）
        if self.current_step == self.total_steps - 1:
            self.cancel_button.config(state=tk.DISABLED)
        else:
            self.cancel_button.config(state=tk.NORMAL)

    def _go_back(self) -> None:
        """前のステップに戻る"""
        if self.current_step > 0:
            self._show_step(self.current_step - 1)

    def _go_next(self) -> None:
        """次のステップに進む"""
        # 現在のステップの検証
        if not self._validate_current_step():
            return

        if self.current_step < self.total_steps - 1:
            self._show_step(self.current_step + 1)
        else:
            # 完了
            self._finish()

    def _cancel(self) -> None:
        """セットアップをキャンセル"""
        result = messagebox.askyesno(
            "セットアップのキャンセル",
            "セットアップをキャンセルしますか？\n\n"
            "後から「⚙️ 設定」タブで設定を行うこともできます。",
            parent=self.window,
        )
        if result:
            self.window.destroy()

    def _validate_current_step(self) -> bool:
        """現在のステップの入力を検証

        Returns:
            検証が成功した場合True
        """
        if self.current_step == 0:  # ようこそ画面
            # 検証不要、常にTrue
            return True

        elif self.current_step == 1:  # 基本設定（年度 + フォルダ）
            year = self.year_var.get().strip()

            if not year:
                messagebox.showerror(
                    "入力エラー", "年度（西暦）を入力してください", parent=self.window
                )
                return False

            # 西暦が4桁の数字かチェック
            if not year.isdigit() or len(year) != 4:
                messagebox.showerror(
                    "入力エラー",
                    "年度は4桁の西暦で入力してください（例: 2026）",
                    parent=self.window,
                )
                return False

            # year_shortは自動計算されるため検証不要

            # フォルダの検証
            folder = self.gdrive_var.get().strip()

            if not folder:
                result = messagebox.askyesno(
                    "確認",
                    "作業フォルダが設定されていません。\n\n"
                    "後から設定することもできますが、続行しますか？",
                    parent=self.window,
                )
                return result

            # パスの検証
            is_valid, error_msg, _ = PathValidator.validate_directory(
                folder, must_exist=False
            )

            if not is_valid:
                messagebox.showerror(
                    "パスエラー", f"無効なパスです:\n{error_msg}", parent=self.window
                )
                return False

        return True

    def _browse_folder(self) -> None:
        """フォルダ選択ダイアログを表示"""
        # 安全な初期ディレクトリを取得
        initial_dir = PathValidator.get_safe_initial_dir(self.gdrive_var.get())

        folder = filedialog.askdirectory(
            parent=self.window, title="作業フォルダを選択", initialdir=str(initial_dir)
        )

        if folder:
            self.gdrive_var.set(folder)
            self._update_folder_status()

    def _update_folder_status(self) -> None:
        """フォルダ状態を更新"""
        folder = self.gdrive_var.get().strip()
        if folder:
            if Path(folder).exists():
                self.folder_status_label.config(
                    text="✓ フォルダが見つかりました", fg="green"
                )
            else:
                self.folder_status_label.config(
                    text="⚠ フォルダが見つかりません", fg="orange"
                )
        else:
            self.folder_status_label.config(text="", fg="gray")

    def _detect_ghostscript_async(self) -> None:
        """Ghostscriptを非同期検出"""
        try:
            gs_path = GhostscriptDetector.detect()
            if gs_path:
                self.gs_var.set(gs_path)
                # ステップ4のUIが作成されている場合のみ更新
                if hasattr(self, "gs_status_label"):
                    self.gs_status_label.config(
                        text="✓ Ghostscriptが見つかりました", fg="green"
                    )
                logger.info(f"Ghostscriptを自動検出: {gs_path}")
            else:
                if hasattr(self, "gs_status_label"):
                    self.gs_status_label.config(
                        text="⚠ Ghostscriptが見つかりませんでした", fg="orange"
                    )
                logger.warning("Ghostscriptが見つかりませんでした")
        except Exception as e:
            logger.error(f"Ghostscript検出エラー: {e}", exc_info=True)
            if hasattr(self, "gs_status_label"):
                self.gs_status_label.config(text="❌ 検出に失敗しました", fg="red")

    def _finish(self) -> None:
        """セットアップを完了して設定を保存"""
        try:
            # 年度を取得し、year_shortを自動計算
            year = self.year_var.get().strip()
            year_short = calculate_year_short(year)

            # 設定を保存
            self.config.set("year", value=year)
            self.config.set("year_short", value=year_short)
            self.config.set(
                "base_paths", "google_drive", value=self.gdrive_var.get().strip()
            )
            self.config.set(
                "base_paths", "local_temp", value=self.local_temp_var.get().strip()
            )

            # v3.5.0: Excelファイル設定は削除（セッション内管理に変更）

            # Ghostscript設定
            if self.gs_var.get():
                self.config.set("ghostscript", "executable", value=self.gs_var.get())
            else:
                self.config.set("ghostscript", "executable", value="")

            self.config.save_config()

            logger.info("初回セットアップが完了しました")

            # ウィンドウを閉じる
            self.window.destroy()

            # 完了コールバック
            if self.on_complete:
                self.on_complete()

        except Exception as e:
            logger.error(f"設定保存エラー: {e}", exc_info=True)
            messagebox.showerror(
                "エラー", f"設定の保存に失敗しました:\n{e}", parent=self.window
            )
