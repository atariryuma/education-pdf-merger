"""
パス検証ユーティリティ

ファイルパスとディレクトリパスの検証を提供
2025年ベストプラクティスに準拠（pathlibベース、Python 3.8+互換）
"""
import logging
import re
import sys
import unicodedata
from pathlib import Path
from typing import Tuple, Optional

from constants import PathConstants

logger = logging.getLogger(__name__)

# モジュールレベルでバージョン判定（起動時に1回だけ実行）
HAS_IS_RELATIVE_TO = sys.version_info >= (3, 9)


def _check_path_security(path: Path, base_resolved: Path) -> bool:
    """
    パスが基準ディレクトリ配下にあるかチェック.

    Args:
        path: チェック対象のパス
        base_resolved: 基準ディレクトリ（解決済み）

    Returns:
        bool: 基準ディレクトリ配下なら True

    Note:
        Python 3.9+ では is_relative_to() を使用、
        Python 3.8 では relative_to() でフォールバック
    """
    if HAS_IS_RELATIVE_TO:
        return path.is_relative_to(base_resolved)
    else:
        # Python 3.8 互換のフォールバック
        try:
            path.relative_to(base_resolved)
            return True
        except ValueError:
            return False


class PathValidationError(Exception):
    """パス検証エラー"""
    pass


class PathValidator:
    """
    パス検証クラス

    pathlibを使用したモダンなパス検証を提供
    セキュリティ（ディレクトリトラバーサル対策）を含む
    """

    # Windows予約名
    WINDOWS_RESERVED_NAMES = (
        'CON', 'PRN', 'AUX', 'NUL',
        'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9',
        'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9'
    )

    @staticmethod
    def normalize_path(path_str: str) -> Path:
        """
        パス文字列を正規化してPathオブジェクトに変換

        Args:
            path_str: パス文字列

        Returns:
            正規化されたPathオブジェクト

        Raises:
            PathValidationError: パスが無効な場合
        """
        if not path_str or not path_str.strip():
            raise PathValidationError("パスが空です")

        # 前後の空白と制御文字を削除
        cleaned = path_str.strip()

        # 改行・タブを削除（途中のスペースは保持）
        cleaned = cleaned.replace('\n', '').replace('\r', '').replace('\t', '')

        try:
            # Pathオブジェクトに変換して正規化
            path = Path(cleaned)

            # 絶対パスに解決（シンボリックリンクも解決）
            resolved_path = path.resolve()

            return resolved_path

        except (OSError, ValueError) as e:
            raise PathValidationError(f"無効なパス形式: {e}") from e

    @staticmethod
    def validate_directory(
        path_str: str,
        must_exist: bool = True,
        base_dir: Optional[Path] = None
    ) -> Tuple[bool, Optional[str], Optional[Path]]:
        """
        ディレクトリパスを検証

        Args:
            path_str: 検証するパス文字列
            must_exist: 存在チェックを行うか
            base_dir: セキュリティチェック用の基準ディレクトリ

        Returns:
            (is_valid, error_message, normalized_path)のタプル
        """
        try:
            # パスを正規化
            path = PathValidator.normalize_path(path_str)

            # セキュリティチェック: ディレクトリトラバーサル対策
            if base_dir is not None:
                base_resolved = base_dir.resolve()
                if not _check_path_security(path, base_resolved):
                    return False, f"許可されていないディレクトリです: {path}", None

            # 存在チェック
            if must_exist:
                if not path.exists():
                    # 親ディレクトリの存在を確認して詳細なエラーメッセージ
                    parent = path.parent
                    if parent.exists():
                        return False, f"ディレクトリが存在しません: {path}\n親ディレクトリは存在します: {parent}", None
                    else:
                        return False, f"ディレクトリが存在しません: {path}", None

                # ディレクトリであることを確認
                if not path.is_dir():
                    return False, f"指定されたパスはディレクトリではありません: {path}", None

            return True, None, path

        except PathValidationError as e:
            return False, str(e), None
        except Exception as e:
            logger.error(f"予期しないパス検証エラー: {e}", exc_info=True)
            return False, f"パス検証中にエラーが発生しました: {e}", None

    @staticmethod
    def validate_file_path(
        path_str: str,
        must_exist: bool = False,
        allowed_extensions: Optional[list] = None
    ) -> Tuple[bool, Optional[str], Optional[Path]]:
        """
        ファイルパスを検証

        Args:
            path_str: 検証するパス文字列
            must_exist: 存在チェックを行うか
            allowed_extensions: 許可する拡張子のリスト（例: ['.pdf', '.txt']）

        Returns:
            (is_valid, error_message, normalized_path)のタプル
        """
        try:
            # パスを正規化
            path = PathValidator.normalize_path(path_str)

            # 拡張子チェック
            if allowed_extensions:
                if path.suffix.lower() not in [ext.lower() for ext in allowed_extensions]:
                    return False, f"許可されていない拡張子です: {path.suffix}\n許可: {', '.join(allowed_extensions)}", None

            # 存在チェック
            if must_exist:
                if not path.exists():
                    parent = path.parent
                    if parent.exists():
                        return False, f"ファイルが存在しません: {path}\n親ディレクトリは存在します: {parent}", None
                    else:
                        return False, f"ファイルが存在しません: {path}", None

                # ファイルであることを確認
                if not path.is_file():
                    return False, f"指定されたパスはファイルではありません: {path}", None
            else:
                # 保存先の場合、親ディレクトリが存在するか確認
                parent = path.parent
                if not parent.exists():
                    return False, f"保存先ディレクトリが存在しません: {parent}", None

            return True, None, path

        except PathValidationError as e:
            return False, str(e), None
        except Exception as e:
            logger.error(f"予期しないファイルパス検証エラー: {e}", exc_info=True)
            return False, f"ファイルパス検証中にエラーが発生しました: {e}", None

    @staticmethod
    def get_safe_initial_dir(path_str: str, fallback: Optional[Path] = None) -> Path:
        """
        ファイルダイアログ用の安全な初期ディレクトリを取得

        Args:
            path_str: ユーザー入力のパス文字列
            fallback: フォールバックディレクトリ（Noneの場合はホームディレクトリ）

        Returns:
            安全な初期ディレクトリPath
        """
        try:
            if path_str and path_str.strip():
                path = PathValidator.normalize_path(path_str)

                # ディレクトリの場合はそのまま
                if path.is_dir():
                    return path

                # ファイルの場合は親ディレクトリ
                if path.parent.exists():
                    return path.parent
        except (OSError, ValueError) as e:
            logger.debug(f"パス正規化失敗: {e}")
            # フォールバックへ

        # フォールバック
        if fallback and fallback.exists():
            return fallback

        # 最終フォールバック: ホームディレクトリ
        return Path.home()

    @staticmethod
    def sanitize_filename(
        filename: str,
        replacement: str = '_',
        max_length: int = 255,
        default_name: str = PathConstants.DEFAULT_FILENAME
    ) -> str:
        """
        ファイル名を安全にサニタイズ（セキュリティベストプラクティス準拠）

        - Null bytes、制御文字、パス区切り文字を除去
        - Windows/Linux/macOSの予約文字を除去
        - Unicode正規化（NFD → NFC）
        - 先頭/末尾のスペース・ドット・アンダースコアを削除
        - 連続する置換文字を統合
        - Windows予約名（CON, PRN等）を回避

        Args:
            filename: サニタイズするファイル名
            replacement: 無効な文字の置換文字
            max_length: 最大文字数（拡張子を含む）
            default_name: 空文字列になった場合のデフォルト名

        Returns:
            サニタイズされたファイル名
        """
        if not filename or not filename.strip():
            return default_name

        # Unicode正規化（NFD → NFC: 結合文字を統合）
        normalized = unicodedata.normalize('NFC', filename)

        # Null bytes と制御文字を削除
        # \x00-\x1f: 制御文字、\x7f: DELETE、\x80-\x9f: 拡張制御文字
        cleaned = ''.join(char for char in normalized if ord(char) >= 0x20 and ord(char) != 0x7f)

        # Windows/Linux/macOSの無効文字を置換
        # < > : " / \ | ? * および \x00-\x1f
        invalid_chars = r'[<>:"/\\|?*\x00-\x1f]'
        cleaned = re.sub(invalid_chars, replacement, cleaned)

        # 連続する置換文字を1つに統合
        if replacement:
            pattern = re.escape(replacement) + r'+'
            cleaned = re.sub(pattern, replacement, cleaned)

        # 先頭と末尾の空白、ドット、アンダースコアを削除
        # Windows: 末尾のスペースとドットは無視される
        cleaned = cleaned.strip(' ._')

        # 空文字列の場合はデフォルト名
        if not cleaned:
            return default_name

        # Windows予約名チェック（クラス変数を使用）
        name_without_ext = Path(cleaned).stem.upper()
        if name_without_ext in PathValidator.WINDOWS_RESERVED_NAMES:
            cleaned = f"{replacement}{cleaned}"

        # 長さ制限（バイト数ではなく文字数）
        if len(cleaned) > max_length:
            # 拡張子を保持しながら切り詰める
            stem = Path(cleaned).stem
            suffix = Path(cleaned).suffix
            max_stem_length = max_length - len(suffix)
            if max_stem_length > 0:
                cleaned = stem[:max_stem_length] + suffix
            else:
                cleaned = cleaned[:max_length]

        return cleaned
