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
from typing import Optional


class StructuredFormatter(logging.Formatter):
    """JSON形式でログを出力するフォーマッター"""

    def format(self, record: logging.LogRecord) -> str:
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
    """機密情報をマスクするフィルター（最適化版）"""

    # 統合パターン（パフォーマンス改善：1つの正規表現に統合）
    _SENSITIVE_PATTERN = re.compile(
        r'(?P<password>password|passwd|pwd)["\']?\s*[:=]\s*["\']?[^"\'}\s,]+'
        r'|(?P<token>token|api[_-]?key|secret|access[_-]?token)["\']?\s*[:=]\s*["\']?[^"\'}\s,]+'
        r'|(?P<credit>\b\d{13,19}\b)'
        r'|(?P<email>\b[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}\b)'
        r'|(?P<phone>\b0\d{1,4}-\d{1,4}-\d{4}\b|\b0\d{9,10}\b)'
        r'|(?P<winpath>(?:C:\\Users\\|/Users/)[^\\\/\s]+)',
        re.I
    )

    @staticmethod
    def _mask_match(match: re.Match) -> str:
        """
        マッチした部分に応じてマスク文字列を返す

        Args:
            match: 正規表現のマッチオブジェクト

        Returns:
            str: マスク処理された文字列
        """
        if match.lastgroup == 'password':
            return 'password=***'
        elif match.lastgroup == 'token':
            return 'token=***'
        elif match.lastgroup == 'credit':
            return '****-****-****-****'
        elif match.lastgroup == 'email':
            # セキュリティ強化: ドメイン部分も完全にマスク
            return '***@***'
        elif match.lastgroup == 'phone':
            return '***-****-****'
        elif match.lastgroup == 'winpath':
            prefix = 'C:\\Users\\' if 'C:' in match.group(0) else '/Users/'
            return f'{prefix}***'
        return '***'

    def filter(self, record: logging.LogRecord) -> bool:
        """ログレコードから機密情報をマスク"""
        try:
            if hasattr(record, 'msg'):
                record.msg = self._SENSITIVE_PATTERN.sub(self._mask_match, str(record.msg))

            if hasattr(record, 'args') and record.args:
                if isinstance(record.args, dict):
                    record.args = {
                        k: self._SENSITIVE_PATTERN.sub(self._mask_match, str(v))
                        for k, v in record.args.items()
                    }
                elif isinstance(record.args, tuple):
                    record.args = tuple(
                        self._SENSITIVE_PATTERN.sub(self._mask_match, str(arg))
                        for arg in record.args
                    )
        except Exception:
            pass  # マスキング失敗時はログ出力を継続

        return True


def setup_logging(log_dir: Optional[str] = None, level: int = logging.INFO, app_name: str = "pdf_merge", use_json: bool = False) -> logging.Logger:
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

    # コンソールハンドラ（UTF-8エンコーディング指定）
    import sys
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(level)
    console_handler.setFormatter(formatter)
    # Windows環境でのUTF-8出力を確保
    if hasattr(sys.stdout, 'reconfigure'):
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except Exception:
            pass  # エンコーディング変更失敗時も継続
    root_logger.addHandler(console_handler)

    return root_logger


def get_logger(name: Optional[str] = None) -> logging.Logger:
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
