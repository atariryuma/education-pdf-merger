"""
設定ファイル読み込みモジュール

このモジュールは、JSON形式の設定ファイルを読み込み、
アプリケーション全体で使用する設定情報を提供します。
デフォルト設定とユーザー設定をマージし、一元管理します。
"""
import copy
import json
import logging
import os
import sys
import time
from typing import Any, Dict, List, Optional, Union, TypeVar

from exceptions import ConfigurationError
from year_utils import calculate_year_short

# ロガーの設定
logger = logging.getLogger(__name__)

# 型変数定義
T = TypeVar('T')


class ConfigLoader:
    """設定ファイルを読み込み、パスを構築するクラス"""

    # デフォルトの設定ファイル名
    DEFAULT_CONFIG_FILENAME = 'config.json'

    def __init__(self, config_path: Optional[str] = None) -> None:
        """
        設定ファイルを読み込む

        Args:
            config_path: 設定ファイルのパス（省略時はこのモジュールと同じディレクトリのconfig.json）
        """
        if config_path is None:
            # PyInstallerでビルドされた場合は実行ファイルと同じディレクトリを使用
            if getattr(sys, 'frozen', False):
                # PyInstallerでビルドされている場合
                module_dir = os.path.dirname(sys.executable)
            else:
                # 通常のPythonスクリプトとして実行されている場合
                module_dir = os.path.dirname(os.path.abspath(__file__))
            config_path = os.path.join(module_dir, self.DEFAULT_CONFIG_FILENAME)

        self.config_path: str = config_path  # デフォルト設定（読み取り専用）

        # ユーザー設定ファイルのパス（AppData内、読み書き可能）
        appdata = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
        user_config_dir = os.path.join(appdata, 'PDFMergeSystem')
        os.makedirs(user_config_dir, exist_ok=True)
        self.user_config_path = os.path.join(user_config_dir, 'user_config.json')

        # ユーザー設定を別途保持（行事名設定などで使用）
        self.user_config: Dict[str, Any] = {}

        self.config: Dict[str, Any] = self._load_config()
        self.year: str = self.config['year']
        # year_shortは自動計算（設定ファイルの値は無視）
        self.year_short: str = calculate_year_short(self.year)

    def _load_config(self) -> Dict[str, Any]:
        """
        設定ファイルを読み込む（デフォルト設定 + ユーザー設定をマージ）

        Returns:
            Dict[str, Any]: 設定辞書

        Raises:
            ConfigurationError: ファイルが見つからない場合またはJSON形式が不正な場合
        """
        # デフォルト設定を読み込み
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
        except FileNotFoundError as e:
            logger.error(f"設定ファイルが見つかりません: {self.config_path}")
            raise ConfigurationError(
                f"設定ファイルが見つかりません: {self.config_path}",
                config_key="config_path",
                original_error=e
            ) from e
        except json.JSONDecodeError as e:
            logger.error(f"設定ファイルのJSON形式が不正です: {e}")
            raise ConfigurationError(
                "設定ファイルのJSON形式が不正です",
                config_key="json_format",
                original_error=e
            ) from e

        # ユーザー設定を読み込んでマージ
        if os.path.exists(self.user_config_path):
            try:
                with open(self.user_config_path, 'r', encoding='utf-8') as f:
                    user_config = json.load(f)
                # インスタンス変数に保存
                self.user_config = user_config
                # ディープマージ
                self._deep_merge(config, user_config)
                logger.info(f"ユーザー設定を読み込みました: {self.user_config_path}")
            except json.JSONDecodeError as e:
                logger.warning(f"ユーザー設定のJSON形式が不正です: {e}")
            except (OSError, PermissionError) as e:
                logger.warning(f"ユーザー設定ファイルの読み込みに失敗しました: {e}")
            except Exception as e:
                logger.warning(f"ユーザー設定の読み込み中に予期しないエラー: {e}")

        # マイグレーション適用
        self._apply_migrations(config)

        return config

    def _apply_migrations(self, config: Dict[str, Any]) -> None:
        """
        設定スキーマのマイグレーション適用

        既存ユーザーの設定に新しいキーを自動追加します。

        Args:
            config: 設定辞書（この辞書が更新される）
        """
    def _deep_merge(self, base: dict, override: dict) -> None:
        """
        辞書を再帰的にマージ（overrideの値でbaseを上書き）

        Args:
            base: ベースとなる辞書（この辞書が更新される）
            override: 上書きする辞書
        """
        for key, value in override.items():
            if key in base and isinstance(base[key], dict) and isinstance(value, dict):
                self._deep_merge(base[key], value)
            else:
                base[key] = copy.deepcopy(value)

    def build_path(self, *parts: str) -> str:
        """
        設定値のプレースホルダーを置換してパスを構築

        Args:
            *parts: パスの各部分

        Returns:
            str: 構築されたパス
        """
        result = []
        for part in parts:
            if isinstance(part, str):
                part = part.replace('{year}', self.year).replace('{year_short}', self.year_short)
            result.append(part)
        return os.path.join(*result)

    def get(self, *keys: str, default: Optional[T] = None) -> Union[Any, T]:
        """
        ネストされた設定値を取得

        Args:
            *keys: 設定のキー（例: 'base_paths', 'google_drive'）
            default: デフォルト値

        Returns:
            設定値（存在しない場合はdefault）

        Examples:
            >>> config.get('year')
            '2026'
            >>> config.get('year_short')
            'R8'
            >>> config.get('base_paths', 'google_drive')
            'G:\\マイドライブ'
            >>> config.get('nonexistent', default='fallback')
            'fallback'
        """
        value: Any = self.config
        for key in keys:
            if isinstance(value, dict) and key in value:
                value = value[key]
            else:
                return default
        return value

    def get_path(self, *path_keys: str, validate: bool = False) -> str:
        """
        設定からパスを取得し、プレースホルダーを置換

        Args:
            *path_keys: パスを構成する設定キーのリスト
            validate: Trueの場合、パスの存在を検証する

        Returns:
            str: 構築されたパス

        Raises:
            ValueError: validate=Trueでパスが存在しない場合
        """
        parts = []
        for key_path in path_keys:
            if isinstance(key_path, str) and '.' in key_path:
                # ドット区切りのキー（例: 'base_paths.google_drive'）
                keys = key_path.split('.')
                value = self.get(*keys)
            else:
                value = key_path
            if value:
                parts.append(value)

        result_path = self.build_path(*parts) if parts else ""

        if validate and result_path and not os.path.exists(result_path):
            raise ValueError(f"パスが存在しません: {result_path}")

        return result_path

    def get_education_plan_path(self) -> str:
        """教育計画のディレクトリパスを取得"""
        # Google Driveパスを取得
        base_path = self.get('base_paths', 'google_drive')
        if not base_path:
            return ""

        return self.build_path(
            base_path,
            self.year,
            self.get('directories', 'education_plan_base'),
            self.get('directories', 'education_plan')
        )

    def get_event_plan_path(self) -> str:
        """行事計画のディレクトリパスを取得"""
        # Google Driveパスを取得
        base_path = self.get('base_paths', 'google_drive')
        if not base_path:
            return ""

        return self.build_path(
            base_path,
            self.year,
            self.get('directories', 'education_plan_base'),
            self.get('directories', 'event_plan')
        )

    def get_temp_dir(self, cleanup_old: bool = False, max_age_hours: int = 24) -> str:
        """
        一時ディレクトリのパスを取得

        Args:
            cleanup_old: 古いファイルをクリーンアップするか
            max_age_hours: 削除対象とするファイルの経過時間（時間）

        Returns:
            str: 一時ディレクトリのパス
        """
        temp_dir = self.get('base_paths', 'local_temp')

        # 設定が空または存在しない場合、デフォルトの一時フォルダを使用
        if not temp_dir:
            appdata = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
            temp_dir = os.path.join(appdata, 'PDFMergeSystem', 'temp')

        if not os.path.exists(temp_dir):
            try:
                os.makedirs(temp_dir, exist_ok=True)
            except PermissionError as e:
                logger.error(f"一時ディレクトリの作成に失敗しました（権限不足）: {temp_dir}")
                raise ConfigurationError(
                    f"一時ディレクトリの作成権限がありません。\n"
                    f"パス: {temp_dir}\n"
                    f"管理者権限で実行するか、別の場所を指定してください。"
                ) from e
            except OSError as e:
                logger.error(f"一時ディレクトリの作成に失敗しました: {temp_dir}, エラー: {e}")
                raise ConfigurationError(
                    f"一時ディレクトリの作成に失敗しました。\n"
                    f"パス: {temp_dir}\n"
                    f"エラー: {e}"
                ) from e

        if cleanup_old:
            self._cleanup_old_temp_files(temp_dir, max_age_hours)

        return temp_dir

    def _cleanup_old_temp_files(self, temp_dir: str, max_age_hours: int) -> None:
        """
        古い一時ファイルをクリーンアップ（再帰的）

        Args:
            temp_dir: 一時ディレクトリのパス
            max_age_hours: 削除対象とするファイルの経過時間（時間）
        """
        current_time = time.time()
        max_age_seconds = max_age_hours * 3600

        try:
            # os.walk()で再帰的にファイルを処理
            for root, dirs, files in os.walk(temp_dir):
                for filename in files:
                    file_path = os.path.join(root, filename)
                    try:
                        file_age = current_time - os.path.getmtime(file_path)
                        if file_age > max_age_seconds:
                            os.remove(file_path)
                            logger.debug(f"古い一時ファイルを削除: {file_path}")
                    except FileNotFoundError:
                        # 既に削除済み（TOCTOU対策）
                        pass
                    except Exception as e:
                        logger.warning(f"一時ファイルの削除に失敗: {file_path} - {e}")

                # 空のディレクトリを削除
                for dirname in dirs:
                    dir_path = os.path.join(root, dirname)
                    try:
                        if not os.listdir(dir_path):  # 空の場合
                            os.rmdir(dir_path)
                            logger.debug(f"空のディレクトリを削除: {dir_path}")
                    except Exception as e:
                        logger.debug(f"ディレクトリ削除スキップ: {dir_path} - {e}")
        except Exception as e:
            logger.warning(f"一時ファイルのクリーンアップに失敗: {e}")

    def set(self, *keys: str, value: Any) -> None:
        """
        ネストされた設定値を設定

        Args:
            *keys: 設定のキー（例: 'base_paths', 'google_drive'）
            value: 設定する値
        """
        if len(keys) == 0:
            return

        current = self.config
        for key in keys[:-1]:
            if key not in current:
                current[key] = {}
            current = current[key]
        current[keys[-1]] = value

    def save_config(self) -> None:
        """
        設定をユーザー設定ファイルに保存

        Raises:
            ConfigurationError: 保存に失敗した場合
        """
        try:
            with open(self.user_config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
            logger.info(f"ユーザー設定を保存しました: {self.user_config_path}")
        except (OSError, PermissionError) as e:
            logger.error(f"ユーザー設定の保存に失敗しました: {e}")
            raise ConfigurationError(
                f"設定ファイルの保存に失敗しました: {self.user_config_path}",
                config_key="save_config"
            ) from e

    def update_year(self, year: str, year_short: Optional[str] = None) -> None:
        """
        年度情報を更新

        Args:
            year: 年度（例: 2026）
            year_short: 年度略称（省略時は自動計算、例: R8）

        Note:
            year_shortは通常省略してください。yearから自動計算されます。
        """
        self.year = year
        # year_shortが明示的に指定されていない場合は自動計算
        self.year_short = year_short if year_short is not None else calculate_year_short(year)
        self.config['year'] = year
        self.config['year_short'] = self.year_short

    def get_event_names(self, category: str) -> List[str]:
        """
        行事名リストを取得（ユーザー設定 > デフォルト値の優先順位）

        Args:
            category: "school_events", "student_council_events", "other_activities"

        Returns:
            行事名のリスト
        """
        # 1. user_config から取得を試みる
        user_event_names = self.user_config.get("excel_event_names", {}).get(category)
        if user_event_names is not None:
            return user_event_names

        # 2. config.json のデフォルト値を使用
        return self.config.get("excel_default_event_names", {}).get(category, [])

    def save_event_names(self, category: str, event_names: List[str]) -> None:
        """
        行事名リストをuser_configに保存

        Args:
            category: "school_events", "student_council_events", "other_activities"
            event_names: 行事名のリスト

        Raises:
            ConfigurationError: 保存に失敗した場合
        """
        if "excel_event_names" not in self.user_config:
            self.user_config["excel_event_names"] = {}

        self.user_config["excel_event_names"][category] = event_names

        # user_config.json に保存
        self._save_user_config()
        logger.info(f"行事名設定を保存しました: {category} ({len(event_names)}件)")

    def reset_event_names(self, category: str) -> bool:
        """
        指定カテゴリの行事名をデフォルトに戻す（user_configから削除）

        Args:
            category: "school_events", "student_council_events", "other_activities"

        Returns:
            bool: 削除が行われた場合True、既にデフォルトの場合False

        Raises:
            ConfigurationError: 保存に失敗した場合
        """
        event_names = self.user_config.get("excel_event_names", {})
        if category not in event_names:
            return False

        del event_names[category]
        # excel_event_namesが空になったら親キーも削除
        if not event_names:
            del self.user_config["excel_event_names"]

        self._save_user_config()
        logger.info(f"行事名設定をデフォルトに戻しました: {category}")
        return True

    def _save_user_config(self) -> None:
        """
        user_config.jsonに保存

        Raises:
            ConfigurationError: 保存に失敗した場合
        """
        try:
            with open(self.user_config_path, 'w', encoding='utf-8') as f:
                json.dump(self.user_config, f, ensure_ascii=False, indent=2)
        except (OSError, PermissionError) as e:
            logger.error(f"ユーザー設定の保存に失敗しました: {e}")
            raise ConfigurationError(
                f"ユーザー設定の保存に失敗しました: {self.user_config_path}",
                config_key="save_user_config",
                original_error=e
            ) from e
