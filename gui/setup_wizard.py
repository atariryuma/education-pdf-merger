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
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Optional, Callable
from pathlib import Path
from datetime import datetime

from config_loader import ConfigLoader
from ghostscript_detector import GhostscriptDetector
from path_validator import PathValidator

logger = logging.getLogger(__name__)


class SetupWizard:
    """初回セットアップウィザード

    5ステップのウィザードで基本設定を完了:
    1. ようこそ画面
    2. 年度設定
    3. 作業フォルダ設定
    4. Ghostscript設定（自動検出）
    5. 完了画面
    """

    def __init__(
        self,
        parent: tk.Tk,
        config: ConfigLoader,
        on_complete: Optional[Callable[[], None]] = None
    ):
        """
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
        self.window.geometry("650x550")
        self.window.resizable(False, False)

        # モーダルダイアログとして設定
        self.window.transient(parent)
        self.window.grab_set()

        # ウィンドウを中央に配置
        self._center_window()

        # 設定値の保持
        self.year_var = tk.StringVar(value=self._get_default_year())
        self.year_short_var = tk.StringVar(value=self._get_default_year_short())
        self.gdrive_var = tk.StringVar(value="")
        self.gs_var = tk.StringVar(value="")
        self.gs_enabled_var = tk.BooleanVar(value=True)

        # 現在のステップ
        self.current_step = 0
        self.total_steps = 5

        # UIコンポーネント
        self.content_frame = None
        self.progress_var = tk.IntVar(value=0)

        # UI構築
        self._create_ui()

        # 最初のステップを表示
        self._show_step(0)

        # Ghostscript自動検出（バックグラウンド）
        self.window.after(100, self._detect_ghostscript_async)

    def _center_window(self) -> None:
        """ウィンドウを画面中央に配置"""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f'{width}x{height}+{x}+{y}')

    def _get_default_year(self) -> str:
        """デフォルトの年度を取得（現在の年度を推定）"""
        now = datetime.now()
        # 4月以降なら来年度、3月以前なら今年度
        if now.month >= 4:
            year = now.year + 1
        else:
            year = now.year

        # 令和年を計算（2019年が令和元年）
        reiwa = year - 2018
        return f"令和{reiwa}年度({year})"

    def _get_default_year_short(self) -> str:
        """デフォルトの年度短縮形を取得"""
        now = datetime.now()
        if now.month >= 4:
            year = now.year + 1
        else:
            year = now.year
        reiwa = year - 2018
        return f"R{reiwa}"

    def _create_ui(self) -> None:
        """UI構築"""
        # ヘッダー
        header_frame = tk.Frame(self.window, bg="#2196F3", height=60)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)

        title_label = tk.Label(
            header_frame,
            text="教育計画PDFマージシステム - 初回セットアップ",
            font=("Yu Gothic UI", 14, "bold"),
            bg="#2196F3",
            fg="white"
        )
        title_label.pack(pady=15)

        # プログレスバー
        progress_frame = tk.Frame(self.window, bg="white", height=40)
        progress_frame.pack(fill=tk.X)
        progress_frame.pack_propagate(False)

        self.progress_label = tk.Label(
            progress_frame,
            text="ステップ 1 / 5",
            font=("Yu Gothic UI", 10),
            bg="white"
        )
        self.progress_label.pack(pady=2)

        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=self.total_steps,
            mode='determinate'
        )
        self.progress_bar.pack(fill=tk.X, padx=20, pady=5)

        # コンテンツエリア
        self.content_frame = tk.Frame(self.window, bg="white")
        self.content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # ボタンフレーム
        button_frame = tk.Frame(self.window, bg="white", height=60)
        button_frame.pack(fill=tk.X, padx=20, pady=10)
        button_frame.pack_propagate(False)

        self.back_button = ttk.Button(
            button_frame,
            text="← 戻る",
            command=self._go_back,
            state=tk.DISABLED
        )
        self.back_button.pack(side=tk.LEFT)

        self.skip_button = ttk.Button(
            button_frame,
            text="スキップ",
            command=self._skip_step,
            state=tk.DISABLED
        )
        self.skip_button.pack(side=tk.LEFT, padx=10)

        self.next_button = ttk.Button(
            button_frame,
            text="次へ →",
            command=self._go_next,
            style="Accent.TButton"
        )
        self.next_button.pack(side=tk.RIGHT)

        self.cancel_button = ttk.Button(
            button_frame,
            text="キャンセル",
            command=self._cancel
        )
        self.cancel_button.pack(side=tk.RIGHT, padx=10)

    def _show_step(self, step: int) -> None:
        """指定されたステップを表示

        Args:
            step: ステップ番号（0-4）
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
            self._show_year_settings()
        elif step == 2:
            self._show_folder_settings()
        elif step == 3:
            self._show_ghostscript_settings()
        elif step == 4:
            self._show_complete()

        # ボタンの状態更新
        self._update_buttons()

    def _show_welcome(self) -> None:
        """ステップ1: ようこそ画面"""
        # タイトル
        title = tk.Label(
            self.content_frame,
            text="ようこそ！",
            font=("Yu Gothic UI", 18, "bold"),
            bg="white"
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
            font=("Yu Gothic UI", 11),
            bg="white",
            justify=tk.LEFT
        )
        desc.pack(pady=20)

        # 機能紹介
        features_frame = tk.LabelFrame(
            self.content_frame,
            text="主な機能",
            font=("Yu Gothic UI", 10, "bold"),
            bg="white",
            relief=tk.GROOVE,
            borderwidth=2
        )
        features_frame.pack(fill=tk.BOTH, expand=True, pady=20)

        features = [
            "📄 Word・Excel・PowerPointのPDF変換",
            "🖼️ 画像ファイルのPDF変換",
            "📝 一太郎文書のPDF変換",
            "📁 フォルダ構造の自動認識",
            "📊 Excel自動転記機能",
            "🗜️ PDF圧縮機能（Ghostscript）"
        ]

        for feature in features:
            label = tk.Label(
                features_frame,
                text=feature,
                font=("Yu Gothic UI", 10),
                bg="white",
                anchor=tk.W
            )
            label.pack(fill=tk.X, padx=20, pady=5)

        # 注意事項
        note = tk.Label(
            self.content_frame,
            text="※ Microsoft Officeがインストールされている必要があります",
            font=("Yu Gothic UI", 9),
            bg="white",
            fg="gray"
        )
        note.pack(pady=10)

    def _show_year_settings(self) -> None:
        """ステップ2: 年度設定"""
        # タイトル
        title = tk.Label(
            self.content_frame,
            text="年度設定",
            font=("Yu Gothic UI", 16, "bold"),
            bg="white"
        )
        title.pack(pady=20)

        # 説明
        desc = tk.Label(
            self.content_frame,
            text="現在の年度を設定してください。\nこの情報はファイル名やフォルダ構造の生成に使用されます。",
            font=("Yu Gothic UI", 10),
            bg="white",
            justify=tk.LEFT
        )
        desc.pack(pady=10)

        # 年度入力
        year_frame = tk.Frame(self.content_frame, bg="white")
        year_frame.pack(pady=20, fill=tk.X)

        year_label = tk.Label(
            year_frame,
            text="年度:",
            font=("Yu Gothic UI", 11),
            bg="white",
            width=15,
            anchor=tk.W
        )
        year_label.pack(side=tk.LEFT, padx=10)

        year_entry = ttk.Entry(
            year_frame,
            textvariable=self.year_var,
            font=("Yu Gothic UI", 11),
            width=30
        )
        year_entry.pack(side=tk.LEFT, padx=10)

        year_hint = tk.Label(
            year_frame,
            text="例: 令和８年度(2026)",
            font=("Yu Gothic UI", 9),
            bg="white",
            fg="gray"
        )
        year_hint.pack(side=tk.LEFT, padx=5)

        # 年度短縮形
        year_short_frame = tk.Frame(self.content_frame, bg="white")
        year_short_frame.pack(pady=10, fill=tk.X)

        year_short_label = tk.Label(
            year_short_frame,
            text="年度（短縮形）:",
            font=("Yu Gothic UI", 11),
            bg="white",
            width=15,
            anchor=tk.W
        )
        year_short_label.pack(side=tk.LEFT, padx=10)

        year_short_entry = ttk.Entry(
            year_short_frame,
            textvariable=self.year_short_var,
            font=("Yu Gothic UI", 11),
            width=30
        )
        year_short_entry.pack(side=tk.LEFT, padx=10)

        year_short_hint = tk.Label(
            year_short_frame,
            text="例: R8",
            font=("Yu Gothic UI", 9),
            bg="white",
            fg="gray"
        )
        year_short_hint.pack(side=tk.LEFT, padx=5)

        # 注意事項
        note_frame = tk.Frame(self.content_frame, bg="#FFF9C4", relief=tk.SOLID, borderwidth=1)
        note_frame.pack(fill=tk.X, pady=20, padx=10)

        note_label = tk.Label(
            note_frame,
            text="💡 この設定は毎年4月に更新してください",
            font=("Yu Gothic UI", 10),
            bg="#FFF9C4",
            anchor=tk.W
        )
        note_label.pack(fill=tk.X, padx=10, pady=10)

    def _show_folder_settings(self) -> None:
        """ステップ3: フォルダ設定"""
        # タイトル
        title = tk.Label(
            self.content_frame,
            text="作業フォルダの設定",
            font=("Yu Gothic UI", 16, "bold"),
            bg="white"
        )
        title.pack(pady=20)

        # 説明
        desc = tk.Label(
            self.content_frame,
            text=(
                "教育計画ファイルが保存されているフォルダを指定してください。\n"
                "Google DriveやOneDriveなどのクラウドストレージも使用できます。"
            ),
            font=("Yu Gothic UI", 10),
            bg="white",
            justify=tk.LEFT
        )
        desc.pack(pady=10)

        # フォルダ選択
        folder_frame = tk.Frame(self.content_frame, bg="white")
        folder_frame.pack(pady=20, fill=tk.X)

        folder_label = tk.Label(
            folder_frame,
            text="作業フォルダ:",
            font=("Yu Gothic UI", 11),
            bg="white",
            width=12,
            anchor=tk.W
        )
        folder_label.pack(side=tk.LEFT, padx=10)

        folder_entry = ttk.Entry(
            folder_frame,
            textvariable=self.gdrive_var,
            font=("Yu Gothic UI", 10),
            width=35
        )
        folder_entry.pack(side=tk.LEFT, padx=5)

        browse_button = ttk.Button(
            folder_frame,
            text="参照...",
            command=self._browse_folder
        )
        browse_button.pack(side=tk.LEFT, padx=5)

        # 状態表示
        self.folder_status_label = tk.Label(
            self.content_frame,
            text="",
            font=("Yu Gothic UI", 9),
            bg="white",
            fg="gray"
        )
        self.folder_status_label.pack(pady=5)

        # ヒント
        hint_frame = tk.Frame(self.content_frame, bg="#E3F2FD", relief=tk.SOLID, borderwidth=1)
        hint_frame.pack(fill=tk.BOTH, expand=True, pady=20, padx=10)

        hint_title = tk.Label(
            hint_frame,
            text="💡 ヒント",
            font=("Yu Gothic UI", 10, "bold"),
            bg="#E3F2FD",
            anchor=tk.W
        )
        hint_title.pack(fill=tk.X, padx=10, pady=5)

        hints = [
            "• Google Driveの「マイドライブ」フォルダを指定できます",
            "• ネットワークドライブも使用可能です",
            "• 後から変更することもできます"
        ]

        for hint in hints:
            hint_label = tk.Label(
                hint_frame,
                text=hint,
                font=("Yu Gothic UI", 9),
                bg="#E3F2FD",
                anchor=tk.W
            )
            hint_label.pack(fill=tk.X, padx=20, pady=2)

    def _show_ghostscript_settings(self) -> None:
        """ステップ4: Ghostscript設定"""
        # タイトル
        title = tk.Label(
            self.content_frame,
            text="PDF圧縮機能の設定",
            font=("Yu Gothic UI", 16, "bold"),
            bg="white"
        )
        title.pack(pady=20)

        # 説明
        desc = tk.Label(
            self.content_frame,
            text=(
                "Ghostscriptを使用すると、生成されたPDFファイルを圧縮できます。\n"
                "自動検出されたパスを確認してください。"
            ),
            font=("Yu Gothic UI", 10),
            bg="white",
            justify=tk.LEFT
        )
        desc.pack(pady=10)

        # Ghostscriptパス
        gs_frame = tk.Frame(self.content_frame, bg="white")
        gs_frame.pack(pady=20, fill=tk.X)

        gs_label = tk.Label(
            gs_frame,
            text="Ghostscript:",
            font=("Yu Gothic UI", 11),
            bg="white",
            width=12,
            anchor=tk.W
        )
        gs_label.pack(side=tk.LEFT, padx=10)

        gs_entry = ttk.Entry(
            gs_frame,
            textvariable=self.gs_var,
            font=("Yu Gothic UI", 9),
            width=40,
            state='readonly'
        )
        gs_entry.pack(side=tk.LEFT, padx=5)

        # 状態表示
        self.gs_status_label = tk.Label(
            self.content_frame,
            text="🔍 検出中...",
            font=("Yu Gothic UI", 10),
            bg="white"
        )
        self.gs_status_label.pack(pady=10)

        # 有効/無効チェックボックス
        enable_checkbox = ttk.Checkbutton(
            self.content_frame,
            text="PDF圧縮機能を使用する",
            variable=self.gs_enabled_var,
            command=self._toggle_ghostscript
        )
        enable_checkbox.pack(pady=10)

        # 注意事項
        note_frame = tk.Frame(self.content_frame, bg="#FFF9C4", relief=tk.SOLID, borderwidth=1)
        note_frame.pack(fill=tk.X, pady=20, padx=10)

        note_text = (
            "ℹ️ Ghostscriptが見つからない場合\n\n"
            "PDF圧縮機能は使用できませんが、その他の機能は正常に動作します。\n"
            "後からGhostscriptをインストールして設定することもできます。"
        )

        note_label = tk.Label(
            note_frame,
            text=note_text,
            font=("Yu Gothic UI", 9),
            bg="#FFF9C4",
            anchor=tk.W,
            justify=tk.LEFT
        )
        note_label.pack(fill=tk.X, padx=10, pady=10)

    def _show_complete(self) -> None:
        """ステップ5: 完了画面"""
        # タイトル
        title = tk.Label(
            self.content_frame,
            text="セットアップ完了！",
            font=("Yu Gothic UI", 18, "bold"),
            bg="white"
        )
        title.pack(pady=30)

        # 成功メッセージ
        message = tk.Label(
            self.content_frame,
            text="基本設定が完了しました。\nアプリケーションを使い始めることができます。",
            font=("Yu Gothic UI", 11),
            bg="white",
            justify=tk.CENTER
        )
        message.pack(pady=20)

        # 設定サマリー
        summary_frame = tk.LabelFrame(
            self.content_frame,
            text="設定内容",
            font=("Yu Gothic UI", 10, "bold"),
            bg="white",
            relief=tk.GROOVE,
            borderwidth=2
        )
        summary_frame.pack(fill=tk.BOTH, expand=True, pady=20, padx=20)

        # 年度
        year_label = tk.Label(
            summary_frame,
            text=f"年度: {self.year_var.get()}",
            font=("Yu Gothic UI", 10),
            bg="white",
            anchor=tk.W
        )
        year_label.pack(fill=tk.X, padx=20, pady=5)

        # 作業フォルダ
        folder_text = self.gdrive_var.get() if self.gdrive_var.get() else "（未設定）"
        folder_label = tk.Label(
            summary_frame,
            text=f"作業フォルダ: {folder_text}",
            font=("Yu Gothic UI", 10),
            bg="white",
            anchor=tk.W
        )
        folder_label.pack(fill=tk.X, padx=20, pady=5)

        # Ghostscript
        gs_text = "有効" if self.gs_enabled_var.get() and self.gs_var.get() else "無効"
        gs_label = tk.Label(
            summary_frame,
            text=f"PDF圧縮機能: {gs_text}",
            font=("Yu Gothic UI", 10),
            bg="white",
            anchor=tk.W
        )
        gs_label.pack(fill=tk.X, padx=20, pady=5)

        # 次のステップ
        next_steps = tk.Label(
            self.content_frame,
            text="設定は「⚙️ 設定」タブからいつでも変更できます。",
            font=("Yu Gothic UI", 9),
            bg="white",
            fg="gray"
        )
        next_steps.pack(pady=20)

    def _update_buttons(self) -> None:
        """ボタンの状態を更新"""
        # 戻るボタン
        if self.current_step == 0:
            self.back_button.config(state=tk.DISABLED)
        else:
            self.back_button.config(state=tk.NORMAL)

        # スキップボタン（Ghostscriptステップのみ）
        if self.current_step == 3:
            self.skip_button.config(state=tk.NORMAL)
        else:
            self.skip_button.config(state=tk.DISABLED)

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

    def _skip_step(self) -> None:
        """現在のステップをスキップ"""
        if self.current_step == 3:  # Ghostscriptステップ
            self.gs_enabled_var.set(False)
            self._go_next()

    def _cancel(self) -> None:
        """セットアップをキャンセル"""
        result = messagebox.askyesno(
            "セットアップのキャンセル",
            "セットアップをキャンセルしますか？\n\n"
            "後から「⚙️ 設定」タブで設定を行うこともできます。",
            parent=self.window
        )
        if result:
            self.window.destroy()

    def _validate_current_step(self) -> bool:
        """現在のステップの入力を検証

        Returns:
            検証が成功した場合True
        """
        if self.current_step == 1:  # 年度設定
            year = self.year_var.get().strip()
            year_short = self.year_short_var.get().strip()

            if not year:
                messagebox.showerror(
                    "入力エラー",
                    "年度を入力してください",
                    parent=self.window
                )
                return False

            if not year_short:
                messagebox.showerror(
                    "入力エラー",
                    "年度（短縮形）を入力してください",
                    parent=self.window
                )
                return False

        elif self.current_step == 2:  # フォルダ設定
            folder = self.gdrive_var.get().strip()

            if not folder:
                result = messagebox.askyesno(
                    "確認",
                    "作業フォルダが設定されていません。\n\n"
                    "後から設定することもできますが、続行しますか？",
                    parent=self.window
                )
                return result

            # パスの検証
            is_valid, error_msg, _ = PathValidator.validate_directory(
                folder,
                must_exist=False
            )

            if not is_valid:
                messagebox.showerror(
                    "パスエラー",
                    f"無効なパスです:\n{error_msg}",
                    parent=self.window
                )
                return False

        return True

    def _browse_folder(self) -> None:
        """フォルダ選択ダイアログを表示"""
        # 安全な初期ディレクトリを取得
        initial_dir = PathValidator.get_safe_initial_dir(self.gdrive_var.get())

        folder = filedialog.askdirectory(
            parent=self.window,
            title="作業フォルダを選択",
            initialdir=str(initial_dir)
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
                    text="✓ フォルダが見つかりました",
                    fg="green"
                )
            else:
                self.folder_status_label.config(
                    text="⚠ フォルダが見つかりません",
                    fg="orange"
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
                if hasattr(self, 'gs_status_label'):
                    self.gs_status_label.config(
                        text="✓ Ghostscriptが見つかりました",
                        fg="green"
                    )
                self.gs_enabled_var.set(True)
                logger.info(f"Ghostscriptを自動検出: {gs_path}")
            else:
                if hasattr(self, 'gs_status_label'):
                    self.gs_status_label.config(
                        text="⚠ Ghostscriptが見つかりませんでした",
                        fg="orange"
                    )
                self.gs_enabled_var.set(False)
                logger.warning("Ghostscriptが見つかりませんでした")
        except Exception as e:
            logger.error(f"Ghostscript検出エラー: {e}", exc_info=True)
            if hasattr(self, 'gs_status_label'):
                self.gs_status_label.config(
                    text="❌ 検出に失敗しました",
                    fg="red"
                )
            self.gs_enabled_var.set(False)

    def _toggle_ghostscript(self) -> None:
        """Ghostscript有効/無効を切り替え"""
        # 現時点では何もしない（チェックボックスの状態のみ保持）
        pass

    def _finish(self) -> None:
        """セットアップを完了して設定を保存"""
        try:
            # 設定を保存
            self.config.set('year', value=self.year_var.get().strip())
            self.config.set('year_short', value=self.year_short_var.get().strip())
            self.config.set('base_paths', 'google_drive', value=self.gdrive_var.get().strip())

            if self.gs_enabled_var.get() and self.gs_var.get():
                self.config.set('ghostscript', 'executable', value=self.gs_var.get())
            else:
                self.config.set('ghostscript', 'executable', value="")

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
                "エラー",
                f"設定の保存に失敗しました:\n{e}",
                parent=self.window
            )
