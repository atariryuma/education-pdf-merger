"""
転送処理ファクトリー

設定に基づいて適切な転送インスタンス（ExcelまたはGoogle Sheets）を生成します。
"""
import logging
import os
from pathlib import Path
from typing import Union, Optional, Callable

from config_loader import ConfigLoader
from update_excel_files import ExcelTransfer
from google_sheets_transfer import GoogleSheetsTransfer
from google_auth_manager import GoogleAuthManager
from exceptions import ConfigurationError, GoogleAuthError

# ロガーの設定
logger = logging.getLogger(__name__)


class HybridTransferFactory:
    """
    Excel/Google Sheets転送処理のファクトリークラス

    設定ファイルの`reference_mode`に基づいて適切な転送インスタンスを生成します。
    """

    @staticmethod
    def create_transfer(
        config: ConfigLoader,
        progress_callback: Optional[Callable[[str], None]] = None,
        cancel_check: Optional[Callable[[], bool]] = None
    ) -> Union[ExcelTransfer, GoogleSheetsTransfer]:
        """
        転送インスタンスを生成

        Args:
            config: 設定ローダー
            progress_callback: 進捗報告用コールバック
            cancel_check: キャンセルチェック用コールバック

        Returns:
            ExcelTransferまたはGoogleSheetsTransferのインスタンス

        Raises:
            ConfigurationError: 設定に問題がある場合
            GoogleAuthError: Google Sheets認証に失敗した場合
        """
        # モードを取得（デフォルト: excel）
        mode = config.get('files', 'reference_mode', default='excel')
        logger.info(f"転送モード: {mode}")

        # ターゲットファイル設定（共通）
        target_filename = config.get('files', 'excel_target')
        target_sheet = config.get('files', 'excel_target_sheet')

        if not target_filename:
            raise ConfigurationError(
                "ターゲットファイルが設定されていません",
                config_key='files.excel_target'
            )

        if mode == 'excel':
            # Excelモード: Excel → Excel転送
            return HybridTransferFactory._create_excel_transfer(
                config, target_filename, target_sheet,
                progress_callback, cancel_check
            )

        elif mode == 'google_sheets':
            # Google Sheetsモード: Google Sheets → Excel転送
            return HybridTransferFactory._create_google_sheets_transfer(
                config, target_filename, target_sheet,
                progress_callback, cancel_check
            )

        else:
            raise ConfigurationError(
                f"無効な参照モード: {mode}\n"
                f"有効な値: 'excel' または 'google_sheets'",
                config_key='files.reference_mode'
            )

    @staticmethod
    def _create_excel_transfer(
        config: ConfigLoader,
        target_filename: str,
        target_sheet: str,
        progress_callback: Optional[Callable[[str], None]],
        cancel_check: Optional[Callable[[], bool]]
    ) -> ExcelTransfer:
        """
        Excel転送インスタンスを生成

        Args:
            config: 設定ローダー
            target_filename: ターゲットファイル名
            target_sheet: ターゲットシート名
            progress_callback: 進捗報告用コールバック
            cancel_check: キャンセルチェック用コールバック

        Returns:
            ExcelTransferインスタンス

        Raises:
            ConfigurationError: 設定に問題がある場合
        """
        ref_filename = config.get('files', 'excel_reference')
        ref_sheet = config.get('files', 'excel_reference_sheet')

        if not ref_filename:
            raise ConfigurationError(
                "参照元Excelファイルが設定されていません",
                config_key='files.excel_reference'
            )

        logger.info(f"Excel転送を作成: {ref_filename} → {target_filename}")

        return ExcelTransfer(
            ref_filename=ref_filename,
            target_filename=target_filename,
            ref_sheet=ref_sheet,
            target_sheet=target_sheet,
            progress_callback=progress_callback,
            cancel_check=cancel_check
        )

    @staticmethod
    def _create_google_sheets_transfer(
        config: ConfigLoader,
        target_filename: str,
        target_sheet: str,
        progress_callback: Optional[Callable[[str], None]],
        cancel_check: Optional[Callable[[], bool]]
    ) -> GoogleSheetsTransfer:
        """
        Google Sheets転送インスタンスを生成

        Args:
            config: 設定ローダー
            target_filename: ターゲットファイル名
            target_sheet: ターゲットシート名
            progress_callback: 進捗報告用コールバック
            cancel_check: キャンセルチェック用コールバック

        Returns:
            GoogleSheetsTransferインスタンス

        Raises:
            ConfigurationError: 設定に問題がある場合
            GoogleAuthError: 認証に失敗した場合
        """
        sheets_url = config.get('files', 'google_sheets_reference_url')
        ref_sheet = config.get('files', 'google_sheets_reference_sheet', default='メインデータ')

        if not sheets_url:
            raise ConfigurationError(
                "Google Sheets URLが設定されていません",
                config_key='files.google_sheets_reference_url'
            )

        logger.info(f"Google Sheets転送を作成: {sheets_url} → {target_filename}")

        # 認証情報ファイルのパスを取得
        credentials_json_path = HybridTransferFactory._get_credentials_path()

        # 認証マネージャーを初期化
        auth_manager = GoogleAuthManager()

        # 認証情報を取得（必要に応じてOAuthフロー実行）
        try:
            if progress_callback:
                progress_callback("Google認証を確認中...")

            credentials = auth_manager.get_credentials(credentials_json_path)
            logger.info("Google認証が完了しました")

        except GoogleAuthError:
            raise
        except Exception as e:
            raise GoogleAuthError(
                "Google認証の取得に失敗しました",
                auth_stage="認証情報取得",
                original_error=e
            ) from e

        return GoogleSheetsTransfer(
            credentials=credentials,
            sheets_url=sheets_url,
            ref_sheet=ref_sheet,
            target_filename=target_filename,
            target_sheet=target_sheet,
            progress_callback=progress_callback,
            cancel_check=cancel_check
        )

    @staticmethod
    def _get_credentials_path() -> str:
        """
        OAuth認証情報ファイル（credentials.json）のパスを取得

        Returns:
            credentials.jsonの絶対パス

        Raises:
            ConfigurationError: credentials.jsonが見つからない場合
        """
        # 1. アプリケーションディレクトリのcredentials.json
        app_dir = Path(__file__).parent
        credentials_path = app_dir / "credentials.json"

        if credentials_path.exists():
            logger.debug(f"認証情報ファイル: {credentials_path}")
            return str(credentials_path)

        # 2. %LOCALAPPDATA%\PDFMergeSystem\credentials.json
        local_appdata = os.environ.get('LOCALAPPDATA', '')
        if local_appdata:
            alt_credentials_path = Path(local_appdata) / "PDFMergeSystem" / "credentials.json"
            if alt_credentials_path.exists():
                logger.debug(f"認証情報ファイル: {alt_credentials_path}")
                return str(alt_credentials_path)

        # 見つからない場合はエラー
        raise ConfigurationError(
            "Google OAuth認証情報ファイル（credentials.json）が見つかりません。\n\n"
            "以下のいずれかの場所に配置してください:\n"
            f"  1. アプリケーションディレクトリ: {app_dir}\n"
            f"  2. ユーザーデータディレクトリ: {Path(local_appdata) / 'PDFMergeSystem'}\n\n"
            "credentials.jsonの取得方法については、ドキュメント（docs/GOOGLE_SHEETS_SETUP.md）を参照してください。",
            config_key="credentials.json"
        )
