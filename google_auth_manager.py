"""
Google OAuth認証マネージャー

Google Sheets APIアクセス用のOAuth 2.0認証を管理
"""
import logging
import os
from pathlib import Path
from typing import Optional

try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
except ImportError as e:
    raise ImportError(
        "Google認証ライブラリがインストールされていません。\n"
        "以下のコマンドを実行してインストールしてください:\n"
        "pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client"
    ) from e

from exceptions import GoogleAuthError

# ロガーの設定
logger = logging.getLogger(__name__)

# Google Sheets APIスコープ（読み取り専用）
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']


class GoogleAuthManager:
    """
    Google OAuth認証マネージャー

    Google Sheets APIへのアクセスに必要な認証情報を管理します。
    """

    def __init__(self, token_dir: Optional[Path] = None) -> None:
        """
        初期化

        Args:
            token_dir: トークン保存ディレクトリ（省略時は%LOCALAPPDATA%/PDFMergeSystem/google_credentials）
        """
        if token_dir is None:
            # デフォルトのトークン保存ディレクトリ
            local_appdata = os.environ.get('LOCALAPPDATA', '')
            if not local_appdata:
                raise GoogleAuthError(
                    "LOCALAPPDATA環境変数が設定されていません",
                    auth_stage="初期化"
                )
            self.token_dir = Path(local_appdata) / "PDFMergeSystem" / "google_credentials"
        else:
            self.token_dir = Path(token_dir)

        # トークンファイルパス
        self.token_path = self.token_dir / "token.json"

        # トークン保存ディレクトリを作成
        try:
            self.token_dir.mkdir(parents=True, exist_ok=True)
            logger.debug(f"トークン保存ディレクトリ: {self.token_dir}")
        except Exception as e:
            raise GoogleAuthError(
                f"トークン保存ディレクトリの作成に失敗しました: {self.token_dir}",
                auth_stage="初期化",
                original_error=e
            ) from e

    def get_credentials(self, credentials_json_path: str) -> Credentials:
        """
        認証情報を取得

        既存のトークンがあれば再利用、なければOAuthフローを実行します。
        トークンが期限切れの場合は自動的に更新します。

        Args:
            credentials_json_path: OAuth 2.0クライアント認証情報ファイル（credentials.json）のパス

        Returns:
            Google API認証情報

        Raises:
            GoogleAuthError: 認証に失敗した場合
        """
        creds = None

        # 既存のトークンを読み込み
        if self.token_path.exists():
            try:
                creds = Credentials.from_authorized_user_file(str(self.token_path), SCOPES)
                logger.debug("既存のトークンを読み込みました")
            except Exception as e:
                logger.warning(f"トークン読み込みに失敗しました: {e}")
                # トークンファイルが破損している可能性があるため削除
                try:
                    self.token_path.unlink()
                    logger.info("破損したトークンファイルを削除しました")
                except Exception:
                    pass

        # トークンが無効または期限切れの場合
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                # トークンを更新
                try:
                    logger.info("トークンを更新しています...")
                    creds.refresh(Request())
                    logger.info("トークンを更新しました")
                except Exception as e:
                    raise GoogleAuthError(
                        "トークンの更新に失敗しました",
                        auth_stage="トークン更新",
                        original_error=e
                    ) from e
            else:
                # 新規認証フロー
                try:
                    logger.info("新規認証フローを開始します...")
                    flow = InstalledAppFlow.from_client_secrets_file(
                        credentials_json_path, SCOPES
                    )
                    creds = flow.run_local_server(port=0)
                    logger.info("認証が完了しました")
                except FileNotFoundError as e:
                    raise GoogleAuthError(
                        f"認証情報ファイルが見つかりません: {credentials_json_path}\n"
                        "Google Cloud Consoleで認証情報を作成し、credentials.jsonとして保存してください。",
                        auth_stage="初回認証",
                        original_error=e
                    ) from e
                except Exception as e:
                    raise GoogleAuthError(
                        "OAuth認証フローに失敗しました",
                        auth_stage="初回認証",
                        original_error=e
                    ) from e

            # トークンを保存
            try:
                with open(self.token_path, 'w', encoding='utf-8') as token_file:
                    token_file.write(creds.to_json())
                logger.debug(f"トークンを保存しました: {self.token_path}")
            except Exception as e:
                logger.warning(f"トークンの保存に失敗しました: {e}")
                # 保存失敗は致命的ではないため、警告のみ

        return creds

    def revoke_credentials(self) -> None:
        """
        認証を解除

        保存されているトークンを削除します。
        次回の認証時には再度OAuthフローが実行されます。

        Raises:
            GoogleAuthError: トークン削除に失敗した場合
        """
        if not self.token_path.exists():
            logger.info("削除するトークンがありません")
            return

        try:
            self.token_path.unlink()
            logger.info("トークンを削除しました")
        except Exception as e:
            raise GoogleAuthError(
                "トークンの削除に失敗しました",
                auth_stage="認証解除",
                original_error=e
            ) from e

    def is_authenticated(self) -> bool:
        """
        認証状態をチェック

        有効なトークンが存在するかを確認します。
        トークンの有効性までは検証しません（実際のAPI呼び出し時に検証されます）。

        Returns:
            トークンファイルが存在すればTrue、存在しなければFalse
        """
        return self.token_path.exists()
