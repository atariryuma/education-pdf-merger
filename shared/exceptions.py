"""
カスタム例外クラス

PDFマージシステム用のカスタム例外を定義
すべての例外クラスで統一されたインターフェースを提供
"""
from typing import Optional


class PDFMergeError(Exception):
    """
    PDF統合エラーの基底クラス

    すべてのカスタム例外の共通インターフェースを定義
    """
    def __init__(
        self,
        message: str,
        *,
        original_error: Optional[Exception] = None,
        **kwargs
    ) -> None:
        """
        Args:
            message: エラーメッセージ
            original_error: 元の例外（チェーン用）
            **kwargs: サブクラス固有の属性
        """
        self.original_error = original_error

        # 元の例外がある場合、メッセージに追加
        if original_error:
            full_message = f"{message} (原因: {type(original_error).__name__}: {original_error})"
        else:
            full_message = message

        super().__init__(full_message)

        # サブクラス固有の属性を保存
        for key, value in kwargs.items():
            setattr(self, key, value)


class PDFConversionError(PDFMergeError):
    """PDF変換エラー"""
    pass


class ConfigurationError(PDFMergeError):
    """設定エラー"""
    def __init__(
        self,
        message: str,
        *,
        config_key: Optional[str] = None,
        original_error: Optional[Exception] = None
    ) -> None:
        """
        Args:
            message: エラーメッセージ
            config_key: 問題のある設定キー
            original_error: 元の例外
        """
        if config_key:
            message = f"設定エラー [{config_key}]: {message}"

        super().__init__(message, original_error=original_error, config_key=config_key)


class ResourceError(PDFMergeError):
    """リソースエラー（COM、ファイル等）"""
    def __init__(
        self,
        message: str,
        *,
        resource_type: str,
        original_error: Optional[Exception] = None
    ) -> None:
        """
        Args:
            message: エラーメッセージ
            resource_type: リソースの種類（例: "Excel COM", "ファイル"）
            original_error: 元の例外
        """
        full_message = f"リソースエラー [{resource_type}]: {message}"
        super().__init__(full_message, original_error=original_error, resource_type=resource_type)


class FileOperationError(PDFMergeError):
    """ファイル操作エラー"""
    def __init__(
        self,
        message: str,
        *,
        file_path: str,
        operation: str,
        original_error: Optional[Exception] = None
    ) -> None:
        """
        Args:
            message: エラーメッセージ
            file_path: 操作対象のファイルパス
            operation: 実行していた操作（例: "読み込み", "書き込み", "削除"）
            original_error: 元の例外
        """
        full_message = f"ファイル{operation}エラー ({file_path}): {message}"
        super().__init__(
            full_message,
            original_error=original_error,
            file_path=file_path,
            operation=operation
        )


class PathNotFoundError(PDFMergeError):
    """パスが見つからないエラー"""
    def __init__(
        self,
        path: str,
        *,
        description: Optional[str] = None,
        original_error: Optional[Exception] = None
    ) -> None:
        """
        Args:
            path: 見つからなかったパス
            description: パスの説明（例: "教育計画フォルダ"）
            original_error: 元の例外
        """
        if description:
            message = f"{description}が見つかりません: {path}"
        else:
            message = f"パスが見つかりません: {path}"

        super().__init__(message, original_error=original_error, path=path, description=description)


class PDFProcessingError(PDFMergeError):
    """PDF処理エラー"""
    def __init__(
        self,
        message: str,
        *,
        operation: str,
        original_error: Optional[Exception] = None
    ) -> None:
        """
        Args:
            message: エラーメッセージ
            operation: 実行していた操作（例: "結合", "分割", "圧縮"）
            original_error: 元の例外
        """
        full_message = f"PDF{operation}エラー: {message}"
        super().__init__(full_message, original_error=original_error, operation=operation)


class ExcelProcessingError(PDFMergeError):
    """Excel処理エラー"""
    def __init__(
        self,
        message: str,
        *,
        file_path: str,
        operation: str,
        original_error: Optional[Exception] = None
    ) -> None:
        """
        Args:
            message: エラーメッセージ
            file_path: 処理対象のファイルパス
            operation: 実行していた操作
            original_error: 元の例外
        """
        full_message = f"Excel{operation}エラー ({file_path}): {message}"
        super().__init__(
            full_message,
            original_error=original_error,
            file_path=file_path,
            operation=operation
        )


class FolderStructureError(PDFMergeError):
    """フォルダ構造分析エラー"""
    def __init__(
        self,
        message: str,
        *,
        directory_path: Optional[str] = None,
        original_error: Optional[Exception] = None
    ) -> None:
        """
        Args:
            message: エラーメッセージ
            directory_path: 分析対象のディレクトリパス
            original_error: 元の例外
        """
        if directory_path:
            message = f"{message} [directory={directory_path}]"

        super().__init__(message, original_error=original_error, directory_path=directory_path)


class CancelledError(PDFMergeError):
    """処理がキャンセルされたことを示す例外"""
    def __init__(self, message: str = "処理がキャンセルされました") -> None:
        super().__init__(message)
