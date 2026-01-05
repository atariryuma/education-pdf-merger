"""
ロギング設定モジュール

統一されたロギングシステムを提供
"""
import logging
from logging.handlers import RotatingFileHandler
import os
import json
import re
from datetime import datetime


class StructuredFormatter(logging.Formatter):
    """JSON形式でログを出力するフォーマッター"""

    def format(self, record):
        log_data = {
            "timestamp": self.formatTime(record, self.datefmt),
            "level": record.levelname,
            "logger": record.name,
            "message": record.getMessage(),
        }

        # 例外情報があれば追加
        if record.exc_info:
            log_data["exception"] = self.formatException(record.exc_info)

        # カスタム属性を追加
        for key, value in record.__dict__.items():
            if key not in ['name', 'msg', 'args', 'created', 'filename', 'funcName',
                           'levelname', 'levelno', 'lineno', 'module', 'msecs',
                           'message', 'pathname', 'process', 'processName',
                           'relativeCreated', 'thread', 'threadName', 'exc_info',
                           'exc_text', 'stack_info']:
                log_data[key] = value

        return json.dumps(log_data, ensure_ascii=False)


class SensitiveDataFilter(logging.Filter):
    """
    機密情報をマスクするフィルター（2025年ベストプラクティス準拠）

    参考:
    - https://betterstack.com/community/guides/logging/sensitive-data/
    - https://dev.to/camillehe1992/mask-sensitive-data-using-python-built-in-logging-module-45fa
    """

    # 包括的な機密データパターン（より強力なセキュリティ）
    SENSITIVE_PATTERNS = [
        # パスワード関連（様々な書式に対応）
        (re.compile(r'password["\']?\s*[:=]\s*["\']?([^"\'}\s,]+)', re.I), 'password=***'),
        (re.compile(r'passwd["\']?\s*[:=]\s*["\']?([^"\'}\s,]+)', re.I), 'passwd=***'),
        (re.compile(r'pwd["\']?\s*[:=]\s*["\']?([^"\'}\s,]+)', re.I), 'pwd=***'),

        # トークン・APIキー関連
        (re.compile(r'token["\']?\s*[:=]\s*["\']?([^"\'}\s,]+)', re.I), 'token=***'),
        (re.compile(r'api[_-]?key["\']?\s*[:=]\s*["\']?([^"\'}\s,]+)', re.I), 'api_key=***'),
        (re.compile(r'secret["\']?\s*[:=]\s*["\']?([^"\'}\s,]+)', re.I), 'secret=***'),
        (re.compile(r'access[_-]?token["\']?\s*[:=]\s*["\']?([^"\'}\s,]+)', re.I), 'access_token=***'),

        # クレジットカード番号（13～19桁）
        (re.compile(r'\b\d{13,19}\b'), '****-****-****-****'),

        # 社会保障番号（米国形式）
        (re.compile(r'\b\d{3}-\d{2}-\d{4}\b'), '***-**-****'),

        # メールアドレス（ユーザー名のみマスク）
        (re.compile(r'\b([a-zA-Z0-9._%+-]+)@([a-zA-Z0-9.-]+\.[a-zA-Z]{2,})\b'), r'***@\2'),

        # 電話番号（日本形式）
        (re.compile(r'\b0\d{1,4}-\d{1,4}-\d{4}\b'), '***-****-****'),
        (re.compile(r'\b0\d{9,10}\b'), '***********'),

        # Windowsパス内のユーザー名（C:\Users\username -> C:\Users\***)
        (re.compile(r'(C:\\Users\\|/Users/)([^\\\/\s]+)', re.I), r'\1***'),
    ]

    def filter(self, record):
        """
        ログレコードから機密情報をマスク

        try-catchでラップして、マスキング失敗時でもログ出力を継続
        （Better Stack推奨のベストプラクティス）
        """
        try:
            if hasattr(record, 'msg'):
                msg = str(record.msg)
                for pattern, replacement in self.SENSITIVE_PATTERNS:
                    msg = pattern.sub(replacement, msg)
                record.msg = msg

            # args属性にも機密情報が含まれる可能性があるため処理
            if hasattr(record, 'args') and record.args:
                if isinstance(record.args, dict):
                    sanitized_args = {}
                    for key, value in record.args.items():
                        str_value = str(value)
                        for pattern, replacement in self.SENSITIVE_PATTERNS:
                            str_value = pattern.sub(replacement, str_value)
                        sanitized_args[key] = str_value
                    record.args = sanitized_args
                elif isinstance(record.args, tuple):
                    sanitized_args = []
                    for arg in record.args:
                        str_arg = str(arg)
                        for pattern, replacement in self.SENSITIVE_PATTERNS:
                            str_arg = pattern.sub(replacement, str_arg)
                        sanitized_args.append(str_arg)
                    record.args = tuple(sanitized_args)
        except Exception:
            # マスキング処理でエラーが発生しても、ログ出力は継続
            # （元のメッセージをそのまま出力）
            pass

        return True


def setup_logging(log_dir=None, level=logging.INFO, app_name="pdf_merge", use_json=False):
    """
    統一ログシステムをセットアップ

    Args:
        log_dir: ログファイルの保存ディレクトリ（省略時はAppData内）
        level: ログレベル（デフォルト: INFO）
        app_name: アプリケーション名（ログファイル名に使用）
        use_json: JSON形式でログを出力するか（デフォルト: False）

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

    # ルートロガーの設定（すべてのモジュールのログをキャッチ）
    root_logger = logging.getLogger()
    root_logger.setLevel(level)

    # 既存のハンドラをクリア（重複防止）
    if root_logger.handlers:
        root_logger.handlers.clear()

    # 機密情報フィルターを追加
    sensitive_filter = SensitiveDataFilter()
    root_logger.addFilter(sensitive_filter)

    # アプリ専用ロガーも設定
    logger = logging.getLogger(app_name)
    logger.setLevel(level)

    # フォーマッター
    if use_json:
        formatter = StructuredFormatter(datefmt='%Y-%m-%d %H:%M:%S')
    else:
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
    root_logger.addHandler(file_handler)

    # コンソールハンドラ
    console_handler = logging.StreamHandler()
    console_handler.setLevel(level)
    console_handler.setFormatter(formatter)
    root_logger.addHandler(console_handler)

    return root_logger


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
