"""
メインアプリケーション

GUIアプリケーションのメインクラス
"""
import json
import logging
import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, messagebox

from config_loader import ConfigLoader
from constants import AppConstants
from gui.styles import WINDOW, FONTS
from gui.tabs.pdf_tab import PDFTab
from gui.tabs.excel_tab import ExcelTab
from gui.tabs.settings_tab import SettingsTab

# ロガーの設定
logger = logging.getLogger(__name__)


def get_app_dir() -> str:
    """アプリケーションのディレクトリを取得（PyInstaller対応）"""
    if getattr(sys, 'frozen', False):
        # PyInstallerでビルドされた場合
        return os.path.dirname(sys.executable)
    else:
        # 通常の実行
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class PDFMergeApp:
    """PDFマージシステムのメインアプリケーションクラス"""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(f"{AppConstants.APP_NAME} v{AppConstants.VERSION}")
        self.root.geometry(WINDOW['geometry'])
        self.root.minsize(WINDOW['min_width'], WINDOW['min_height'])

        # スレッドセーフティのためのロック
        self.config_lock = threading.RLock()

        # 最後の設定を保存するファイル（AppData内に保存）
        appdata = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
        settings_dir = os.path.join(appdata, 'PDFMergeSystem')
        if not os.path.exists(settings_dir):
            try:
                os.makedirs(settings_dir, exist_ok=True)
            except (PermissionError, OSError) as e:
                # 設定ディレクトリ作成失敗時はエラーダイアログを表示して終了
                messagebox.showerror(
                    "起動エラー",
                    f"設定ディレクトリの作成に失敗しました。\n\n"
                    f"パス: {settings_dir}\n"
                    f"エラー: {e}\n\n"
                    f"管理者権限で実行するか、アプリケーションを再インストールしてください。"
                )
                raise
        self.last_settings_file = os.path.join(settings_dir, ".last_settings.json")

        # 設定の読み込み
        try:
            self.config = ConfigLoader()
        except Exception as e:
            messagebox.showerror(
                "設定エラー",
                f"設定ファイルの読み込みに失敗しました。\n\n詳細: {e}\n\nconfig.jsonを確認してください。"
            )
            self.root.destroy()
            return

        # 最後の設定を読み込み
        last_settings = self._load_last_settings()

        # 変数の初期化
        self._init_variables(last_settings)

        # UIを構築
        self._create_ui()

        # キーボードショートカット設定
        self._setup_keyboard_shortcuts()

        # 終了時の処理
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)

        # 初回起動チェック（パス未設定の場合はガイダンスを表示）
        self.root.after(500, self._check_initial_setup)

    def _init_variables(self, last_settings: dict) -> None:
        """変数を初期化"""
        # PDF統合タブ用（入力・出力パスは空で開始、計画種別のみlast_settingsから復元）
        # ネットワークパスはアクセスに時間がかかるため、起動時は空にする
        self.input_dir_var = tk.StringVar(value="")
        self.output_file_var = tk.StringVar(value="")
        self.plan_type_var = tk.StringVar(value=last_settings.get('plan_type', 'education'))

        # 設定タブ用
        self.year_var = tk.StringVar(value=self.config.year)
        self.year_short_var = tk.StringVar(value=self.config.year_short)

        # Google Driveパスのデフォルト値を設定
        gdrive_path = self.config.get('base_paths', 'google_drive') or ""

        self.gdrive_var = tk.StringVar(value=gdrive_path)

        # 一時フォルダ：空の場合はデフォルトパスを設定してconfig.jsonに保存
        temp_path = self.config.get('base_paths', 'local_temp')
        if not temp_path:
            appdata = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
            temp_path = os.path.join(appdata, 'PDFMergeSystem', 'temp')
            self.config.set('base_paths', 'local_temp', value=temp_path)
            self.config.save_config()
        self.temp_var = tk.StringVar(value=temp_path)

        self.gs_var = tk.StringVar(value=self.config.get('ghostscript', 'executable'))
        self.excel_ref_var = tk.StringVar(value=self.config.get('files', 'excel_reference'))
        self.excel_target_var = tk.StringVar(value=self.config.get('files', 'excel_target'))

    def _load_last_settings(self) -> dict:
        """最後の設定を読み込み"""
        try:
            if os.path.exists(self.last_settings_file):
                with open(self.last_settings_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except (IOError, json.JSONDecodeError, OSError):
            pass
        return {}

    def _save_last_settings(self) -> None:
        """最後の設定を保存"""
        try:
            settings = {
                'input_dir': self.input_dir_var.get(),
                'output_file': self.output_file_var.get(),
                'plan_type': self.plan_type_var.get()
            }
            with open(self.last_settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
        except (IOError, OSError):
            pass

    def _setup_keyboard_shortcuts(self) -> None:
        """キーボードショートカット設定"""
        self.root.bind('<Control-s>', lambda e: self._save_settings())
        self.root.bind('<Control-r>', lambda e: self._reload_settings())
        self.root.bind('<Control-q>', lambda e: self._on_closing())
        self.root.bind('<F5>', self._handle_f5)

    def _handle_f5(self, event) -> str:
        """F5キーの処理"""
        if not hasattr(self, 'notebook') or self.notebook is None:
            return 'break'
        try:
            current_tab = self.notebook.index(self.notebook.select())
            if current_tab == 1:  # Excel処理タブ
                self.excel_tab.check_files_status()
        except Exception as e:
            logger.debug(f"F5キー処理でエラー: {e}")
        return 'break'

    def _on_closing(self) -> None:
        """終了時の処理"""
        self._save_last_settings()
        self.root.destroy()

    def _create_ui(self) -> None:
        """UIを構築"""
        # メニューバー
        self._create_menu()

        # ステータスバー
        self.status_bar = tk.Label(
            self.root,
            text="準備完了",
            relief=tk.SUNKEN,
            anchor="w",
            font=FONTS['small']
        )
        self.status_bar.pack(side="bottom", fill="x")

        # タブコントロール
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # 各タブを作成
        self.pdf_tab = PDFTab(
            self.notebook, self.config, self.status_bar,
            self.input_dir_var, self.output_file_var, self.plan_type_var
        )

        self.excel_tab = ExcelTab(self.notebook, self.config, self.status_bar)

        # File Management タブは未実装のため非表示
        # self.file_tab = FileTab(self.notebook, self.config, self.status_bar)

        self.settings_tab = SettingsTab(
            self.notebook, self.config, self.status_bar,
            self.year_var, self.year_short_var,
            self.gdrive_var, self.temp_var, self.gs_var,
            self.excel_ref_var, self.excel_target_var,
            on_reload=self._reload_settings
        )

    def _create_menu(self) -> None:
        """メニューバーを作成"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # ファイルメニュー
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="ファイル", menu=file_menu)
        file_menu.add_command(label="設定を保存 (Ctrl+S)", command=self._save_settings)
        file_menu.add_command(label="設定を再読み込み (Ctrl+R)", command=self._reload_settings)
        file_menu.add_separator()
        file_menu.add_command(label="終了 (Ctrl+Q)", command=self._on_closing)

        # ヘルプメニュー
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="ヘルプ", menu=help_menu)
        help_menu.add_command(label="キーボードショートカット", command=self._show_shortcuts)
        help_menu.add_command(label="バージョン情報", command=self._show_version)

    def _save_settings(self) -> None:
        """設定を保存"""
        self.settings_tab.save_settings()

    def _reload_settings(self) -> None:
        """設定を再読み込み"""
        try:
            self.config = ConfigLoader()
            # UI変数を更新
            self.year_var.set(self.config.year)
            self.year_short_var.set(self.config.year_short)
            self.gdrive_var.set(self.config.get('base_paths', 'google_drive'))
            self.temp_var.set(self.config.get('base_paths', 'local_temp'))
            self.gs_var.set(self.config.get('ghostscript', 'executable'))
            self.excel_ref_var.set(self.config.get('files', 'excel_reference'))
            self.excel_target_var.set(self.config.get('files', 'excel_target'))

            # Excelタブのラベルも更新
            self.excel_tab.update_labels()

            # タブのconfigを更新
            self.pdf_tab.config = self.config
            self.excel_tab.config = self.config
            # self.file_tab.config = self.config  # File Tab は非表示
            self.settings_tab.config = self.config

            self._update_status("設定を再読み込みしました")
            messagebox.showinfo("再読み込み完了", "設定を再読み込みしました")
        except Exception as e:
            messagebox.showerror("読み込みエラー", f"設定の再読み込みに失敗しました。\n\n詳細: {e}")

    def _update_status(self, message: str) -> None:
        """ステータスバーを更新"""
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.status_bar.config(text=f"[{timestamp}] {message}")

    def _show_shortcuts(self) -> None:
        """キーボードショートカットを表示"""
        shortcuts = """
キーボードショートカット一覧

【全般】
Ctrl+S : 設定を保存
Ctrl+R : 設定を再読み込み
Ctrl+Q : アプリを終了

【Excel処理タブ】
F5 : ファイル状態を確認
        """
        messagebox.showinfo("キーボードショートカット", shortcuts.strip())

    def _show_version(self) -> None:
        """バージョン情報を表示"""
        version_info = f"""
{AppConstants.APP_NAME}

バージョン: {AppConstants.VERSION}
作成日: 2025年

【機能】
• PDF統合（教育計画・行事計画）
• Excel自動更新
• ファイル名整理
• 不要シート削除
        """
        messagebox.showinfo("バージョン情報", version_info.strip())

    def _check_initial_setup(self) -> None:
        """初回起動時の設定チェックとセットアップウィザード表示"""
        try:
            from config_validator import ConfigValidator
            from gui.setup_wizard import SetupWizard

            # 設定を検証
            validator = ConfigValidator(self.config)
            is_valid, results = validator.validate_all()

            # エラーレベルの問題がある場合、セットアップウィザードを表示
            if validator.has_errors():
                logger.info("必須設定が不足しています。セットアップウィザードを起動します。")

                def on_wizard_complete():
                    """ウィザード完了時のコールバック"""
                    # 設定を再読み込み
                    self.config = ConfigLoader()
                    self._init_variables({})

                    # UIを更新（すべてのタブに反映）
                    self.year_var.set(self.config.year)
                    self.year_short_var.set(self.config.year_short)
                    self.gdrive_var.set(self.config.get('base_paths', 'google_drive') or "")
                    self.temp_var.set(self.config.get('base_paths', 'local_temp') or "")
                    self.gs_var.set(self.config.get('ghostscript', 'executable') or "")
                    self.excel_ref_var.set(self.config.get('files', 'excel_reference') or "")
                    self.excel_target_var.set(self.config.get('files', 'excel_target') or "")

                    # 各タブのconfigも更新
                    if hasattr(self, 'pdf_tab'):
                        self.pdf_tab.config = self.config
                    if hasattr(self, 'excel_tab'):
                        self.excel_tab.config = self.config
                    if hasattr(self, 'settings_tab'):
                        self.settings_tab.config = self.config
                    if hasattr(self, 'file_tab'):
                        self.file_tab.config = self.config

                    messagebox.showinfo(
                        "セットアップ完了",
                        "設定が完了しました！\n\nアプリケーションを使い始めることができます。",
                        parent=self.root
                    )
                    logger.info("セットアップウィザードが完了しました")

                # セットアップウィザードを表示
                SetupWizard(self.root, self.config, on_complete=on_wizard_complete)

            elif validator.has_warnings():
                # 警告のみの場合は通知のみ
                logger.warning(f"設定に警告があります: {validator.get_summary()}")

        except Exception as e:
            logger.error(f"初回セットアップチェックエラー: {e}", exc_info=True)



def main() -> None:
    """メイン関数"""
    root = tk.Tk()
    app = PDFMergeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
