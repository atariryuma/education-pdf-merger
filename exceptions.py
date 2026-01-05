"""
カスタム例外クラス

PDFマージシステム用のカスタム例外を定義
"""
from datetime import datetime
from typing import Optional, Any


class PDFMergeError(Exception):
    """PDF統合エラーの基底クラス"""

    def __init__(self, message: str):
        super().__init__(message)
        self.timestamp = datetime.now()
        self.context = {}

    def add_context(self, key: str, value: Any):
        """エラーコンテキストに情報を追加"""
        self.context[key] = value
        return self

    def __str__(self):
        base_msg = super().__str__()
        if self.context:
            ctx_str = ", ".join(f"{k}={v}" for k, v in self.context.items())
            return f"{base_msg} [{ctx_str}]"
        return base_msg


class PDFConversionError(PDFMergeError):
    """PDF変換エラー"""

    def __init__(self, message_or_file_path: str, operation: str = None, original_error: Exception = None):
        """
        Args:
            message_or_file_path: エラーメッセージまたは変換に失敗したファイルのパス
            operation: 実行していた操作（例: "Word変換", "Excel変換"）。省略時はmessage_or_file_pathがメッセージとして扱われる
            original_error: 元の例外オブジェクト
        """
        if operation is None:
            # シンプルなメッセージのみの場合（後方互換性）
            self.file_path = None
            self.operation = None
            self.original_error = original_error
            self.original_format = None
            message = message_or_file_path
        else:
            # 詳細な情報付きの場合
            self.file_path = message_or_file_path
            self.operation = operation
            self.original_error = original_error
            # ファイル拡張子から元の形式を取得
            import os
            self.original_format = os.path.splitext(message_or_file_path)[1].lower() if message_or_file_path else None
            message = f"{operation}失敗: {message_or_file_path}"
            if original_error:
                message += f" - {original_error}"
        super().__init__(message)
        # タイムスタンプと形式情報をコンテキストに追加
        if self.original_format:
            self.add_context('format', self.original_format)
        if self.file_path:
            self.add_context('file_path', self.file_path)


class ConfigurationError(PDFMergeError):
    """設定エラー"""

    def __init__(self, message, config_key=None):
        """
        Args:
            message: エラーメッセージ
            config_key: 問題のある設定キー
        """
        self.config_key = config_key
        if config_key:
            message = f"設定エラー [{config_key}]: {message}"
        super().__init__(message)


class ResourceError(PDFMergeError):
    """リソースエラー（COM、ファイル等）"""

    def __init__(self, resource_type, message, original_error=None):
        """
        Args:
            resource_type: リソースの種類（例: "Excel COM", "ファイル"）
            message: エラーメッセージ
            original_error: 元の例外オブジェクト
        """
        self.resource_type = resource_type
        self.original_error = original_error
        full_message = f"リソースエラー [{resource_type}]: {message}"
        if original_error:
            full_message += f" - {original_error}"
        super().__init__(full_message)


class FileOperationError(PDFMergeError):
    """ファイル操作エラー"""

    def __init__(self, file_path, operation, original_error=None):
        """
        Args:
            file_path: 操作対象のファイルパス
            operation: 実行していた操作（例: "読み込み", "書き込み", "削除"）
            original_error: 元の例外オブジェクト
        """
        self.file_path = file_path
        self.operation = operation
        self.original_error = original_error
        message = f"ファイル{operation}エラー: {file_path}"
        if original_error:
            message += f" - {original_error}"
        super().__init__(message)


class PathNotFoundError(PDFMergeError):
    """パスが見つからないエラー"""

    def __init__(self, path, description=None):
        """
        Args:
            path: 見つからなかったパス
            description: パスの説明（例: "教育計画フォルダ"）
        """
        self.path = path
        self.description = description
        message = f"パスが見つかりません: {path}"
        if description:
            message = f"{description}が見つかりません: {path}"
        super().__init__(message)


class PDFProcessingError(PDFMergeError):
    """PDF処理エラー"""

    def __init__(self, operation, message, original_error=None):
        """
        Args:
            operation: 実行していた操作（例: "結合", "分割", "圧縮"）
            message: エラーメッセージ
            original_error: 元の例外オブジェクト
        """
        self.operation = operation
        self.original_error = original_error
        full_message = f"PDF{operation}エラー: {message}"
        if original_error:
            full_message += f" - {original_error}"
        super().__init__(full_message)


class ExcelProcessingError(PDFMergeError):
    """Excel処理エラー"""

    def __init__(self, file_path, operation, original_error=None):
        """
        Args:
            file_path: 処理対象のファイルパス
            operation: 実行していた操作
            original_error: 元の例外オブジェクト
        """
        self.file_path = file_path
        self.operation = operation
        self.original_error = original_error
        message = f"Excel{operation}エラー: {file_path}"
        if original_error:
            message += f" - {original_error}"
        super().__init__(message)


class FolderStructureError(PDFMergeError):
    """フォルダ構造分析エラー"""

    def __init__(self, message: str, directory_path: str = None):
        """
        Args:
            message: エラーメッセージ
            directory_path: 分析対象のディレクトリパス
        """
        self.directory_path = directory_path
        super().__init__(message)
        if directory_path:
            self.add_context('directory', directory_path)


class CancelledError(PDFMergeError):
    """処理がキャンセルされたことを示す例外"""
    pass
