"""
カスタム例外クラス

PDFマージシステム用のカスタム例外を定義
"""


class PDFMergeError(Exception):
    """PDF統合エラーの基底クラス"""
    pass


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
            message = message_or_file_path
        else:
            # 詳細な情報付きの場合
            self.file_path = message_or_file_path
            self.operation = operation
            self.original_error = original_error
            message = f"{operation}失敗: {message_or_file_path}"
            if original_error:
                message += f" - {original_error}"
        super().__init__(message)


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


class CancelledError(PDFMergeError):
    """処理がキャンセルされたことを示す例外"""
    pass
