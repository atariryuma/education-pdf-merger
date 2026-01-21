"""
教育計画PDFマージシステム - エントリーポイント

PyInstallerでのビルド用メインスクリプト

完全型ヒント、例外安全性、コード品質100点を目指した実装
"""
import sys
import os
import logging
from typing import NoReturn, Type, Optional
import traceback
from types import TracebackType

# COMスレッディングモデルの設定（pywinautoとtkinter.filedialogの競合を解決）
# 参照: https://github.com/pywinauto/pywinauto/issues/517
# 参照: https://bugs.python.org/issue34029
# COINIT_APARTMENTTHREADED (STA) を使用してtkinter.filedialogのフリーズを防止
sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED

# tkinter の可用性チェック（モジュールトップで実行）
try:
    from tkinter import messagebox
    HAS_TKINTER = True
except ImportError:
    HAS_TKINTER = False

# プロジェクトルートをパスに追加
if getattr(sys, 'frozen', False):
    # PyInstallerでビルドされた場合
    application_path = os.path.dirname(sys.executable)
else:
    # 通常の実行
    application_path = os.path.dirname(os.path.abspath(__file__))

sys.path.insert(0, application_path)

# ロギングの設定（環境変数 LOG_LEVEL で制御可能）
from logging_config import setup_logging  # noqa: E402

# ログレベルマッピング（定数として定義）
LOG_LEVEL_MAP = {
    'DEBUG': logging.DEBUG,
    'INFO': logging.INFO,
    'WARNING': logging.WARNING,
    'ERROR': logging.ERROR,
    'CRITICAL': logging.CRITICAL
}

# 環境変数からログレベルを取得
log_level_str = os.environ.get('LOG_LEVEL', 'INFO').upper()
log_level = LOG_LEVEL_MAP.get(log_level_str, logging.INFO)

# ロギングシステムを初期化（この後に logger を取得する）
setup_logging(level=log_level)

# ロガーを取得（setup_logging の後に実行することが重要）
logger = logging.getLogger(__name__)


def global_exception_handler(
    exc_type: Type[BaseException],
    exc_value: BaseException,
    exc_traceback: Optional[TracebackType]
) -> None:
    """
    キャッチされなかった例外を処理するグローバルハンドラ.

    Args:
        exc_type: 例外の型
        exc_value: 例外のインスタンス
        exc_traceback: トレースバック情報

    Note:
        KeyboardInterrupt と SystemExit は標準処理に委譲する
    """
    # KeyboardInterrupt と SystemExit は標準処理
    if issubclass(exc_type, (KeyboardInterrupt, SystemExit)):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return

    # エラーログ出力
    error_msg = ''.join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    logger.critical(f"未処理の例外が発生しました:\n{error_msg}")

    # ユーザーへ通知（GUIが利用可能な場合のみ）
    if HAS_TKINTER:
        try:
            messagebox.showerror(
                "予期しないエラー",
                f"アプリケーションエラーが発生しました。\n\n"
                f"エラー: {exc_type.__name__}: {exc_value}\n\n"
                f"詳細はログファイルを確認してください。"
            )
        except Exception as gui_error:
            # GUIが壊れている場合は諦める
            logger.debug(f"GUI通知失敗: {gui_error}")

    # stderr にも出力（GUI が使えない環境用）
    print(f"Critical Error: {exc_type.__name__}: {exc_value}", file=sys.stderr)


def main() -> NoReturn:
    """
    メインアプリケーションを起動.

    Raises:
        Exception: アプリケーション起動失敗時
    """
    # グローバル例外ハンドラを設定
    sys.excepthook = global_exception_handler

    logger.info("アプリケーションを起動します")

    try:
        from gui.app import main as app_main
        app_main()
    except Exception as e:
        logger.critical(f"アプリケーション起動に失敗しました: {e}", exc_info=True)
        raise


if __name__ == "__main__":
    main()
