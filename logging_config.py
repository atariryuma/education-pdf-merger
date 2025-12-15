"""
ロギング設定モジュール

統一されたロギングシステムを提供
"""
import logging
from logging.handlers import RotatingFileHandler
import os
from datetime import datetime


def setup_logging(log_dir=None, level=logging.INFO, app_name="pdf_merge"):
    """
    統一ログシステムをセットアップ

    Args:
        log_dir: ログファイルの保存ディレクトリ（省略時はAppData内）
        level: ログレベル（デフォルト: INFO）
        app_name: アプリケーション名（ログファイル名に使用）

    Returns:
        logging.Logger: 設定済みのロガー
    """
    # ログディレクトリの決定（AppData内に作成）
    if log_dir is None:
        appdata = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
        log_dir = os.path.join(appdata, 'PDFMergeSystem', 'logs')

    # ログディレクトリの作成
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    # ログファイル名（日付入り）
    log_file = os.path.join(log_dir, f"{app_name}_{datetime.now():%Y%m%d}.log")

    # ルートロガーの設定
    logger = logging.getLogger(app_name)
    logger.setLevel(level)

    # 既存のハンドラをクリア（重複防止）
    if logger.handlers:
        logger.handlers.clear()

    # フォーマッター
    formatter = logging.Formatter(
        '%(asctime)s [%(levelname)s] %(name)s: %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    # ファイルハンドラ（ローテーション付き: 5MB x 5ファイル）
    file_handler = RotatingFileHandler(
        log_file,
        maxBytes=5*1024*1024,
        backupCount=5,
        encoding='utf-8'
    )
    file_handler.setLevel(level)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # コンソールハンドラ
    console_handler = logging.StreamHandler()
    console_handler.setLevel(level)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    return logger


def get_logger(name=None):
    """
    名前付きロガーを取得

    Args:
        name: ロガー名（省略時は'pdf_merge'）

    Returns:
        logging.Logger: ロガーインスタンス
    """
    if name is None:
        name = "pdf_merge"
    return logging.getLogger(name)
