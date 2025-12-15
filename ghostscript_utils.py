"""
Ghostscriptユーティリティモジュール

Ghostscriptの検出、インストール確認、パス設定を行う
"""
import os
import subprocess
import logging
import winreg
from typing import Optional, List

logger = logging.getLogger(__name__)


class GhostscriptManager:
    """Ghostscriptの検出と管理を行うクラス"""

    # 検索対象のディレクトリ
    SEARCH_PATHS = [
        r"C:\Program Files\gs",
        r"C:\Program Files (x86)\gs",
        r"C:\gs",
    ]

    # 実行ファイル名（優先順）
    EXECUTABLE_NAMES = ["gswin64c.exe", "gswin32c.exe"]

    # ダウンロードURL
    DOWNLOAD_URL = "https://ghostscript.com/releases/gsdnld.html"

    @classmethod
    def find_ghostscript(cls) -> Optional[str]:
        """
        システムにインストールされているGhostscriptを検索

        Returns:
            str: Ghostscript実行ファイルのパス（見つからない場合はNone）
        """
        # 1. レジストリから検索
        gs_path = cls._find_from_registry()
        if gs_path:
            logger.info(f"レジストリからGhostscriptを検出: {gs_path}")
            return gs_path

        # 2. 既知のパスから検索
        gs_path = cls._find_from_known_paths()
        if gs_path:
            logger.info(f"既知のパスからGhostscriptを検出: {gs_path}")
            return gs_path

        # 3. PATH環境変数から検索
        gs_path = cls._find_from_path_env()
        if gs_path:
            logger.info(f"PATH環境変数からGhostscriptを検出: {gs_path}")
            return gs_path

        logger.warning("Ghostscriptが見つかりませんでした")
        return None

    @classmethod
    def _find_from_registry(cls) -> Optional[str]:
        """レジストリからGhostscriptのパスを検索"""
        registry_paths = [
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\GPL Ghostscript"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\GPL Ghostscript"),
            (winreg.HKEY_CURRENT_USER, r"SOFTWARE\GPL Ghostscript"),
        ]

        for hkey, subkey in registry_paths:
            try:
                with winreg.OpenKey(hkey, subkey) as key:
                    # サブキー（バージョン番号）を列挙
                    versions = []
                    i = 0
                    while True:
                        try:
                            version = winreg.EnumKey(key, i)
                            versions.append(version)
                            i += 1
                        except OSError:
                            break

                    # 最新バージョンを選択
                    if versions:
                        versions.sort(reverse=True)
                        latest_version = versions[0]
                        with winreg.OpenKey(key, latest_version) as version_key:
                            gs_dir, _ = winreg.QueryValueEx(version_key, "GS_DLL")
                            # DLLパスからbinディレクトリを推測
                            bin_dir = os.path.dirname(gs_dir)
                            for exe_name in cls.EXECUTABLE_NAMES:
                                exe_path = os.path.join(bin_dir, exe_name)
                                if os.path.exists(exe_path):
                                    return exe_path
            except OSError:
                continue

        return None

    @classmethod
    def _find_from_known_paths(cls) -> Optional[str]:
        """既知のインストールパスからGhostscriptを検索"""
        for base_path in cls.SEARCH_PATHS:
            if not os.path.exists(base_path):
                continue

            # gsXX.XX.X のようなバージョンフォルダを検索
            try:
                subdirs = [d for d in os.listdir(base_path)
                          if os.path.isdir(os.path.join(base_path, d)) and d.startswith("gs")]
                # バージョン順にソート（降順）
                subdirs.sort(reverse=True)

                for subdir in subdirs:
                    bin_dir = os.path.join(base_path, subdir, "bin")
                    if os.path.exists(bin_dir):
                        for exe_name in cls.EXECUTABLE_NAMES:
                            exe_path = os.path.join(bin_dir, exe_name)
                            if os.path.exists(exe_path):
                                return exe_path
            except OSError:
                continue

        return None

    @classmethod
    def _find_from_path_env(cls) -> Optional[str]:
        """PATH環境変数からGhostscriptを検索"""
        for exe_name in cls.EXECUTABLE_NAMES:
            try:
                result = subprocess.run(
                    ["where", exe_name],
                    capture_output=True,
                    text=True,
                    timeout=5
                )
                if result.returncode == 0:
                    paths = result.stdout.strip().split("\n")
                    if paths and os.path.exists(paths[0]):
                        return paths[0]
            except Exception:
                continue

        return None

    @classmethod
    def verify_ghostscript(cls, gs_path: str) -> bool:
        """
        Ghostscriptが正常に動作するか確認

        Args:
            gs_path: Ghostscript実行ファイルのパス

        Returns:
            bool: 正常に動作する場合True
        """
        if not gs_path or not os.path.exists(gs_path):
            return False

        try:
            result = subprocess.run(
                [gs_path, "--version"],
                capture_output=True,
                text=True,
                timeout=10
            )
            if result.returncode == 0:
                version = result.stdout.strip()
                logger.info(f"Ghostscript バージョン: {version}")
                return True
        except Exception as e:
            logger.warning(f"Ghostscript検証エラー: {e}")

        return False

    @classmethod
    def get_install_instructions(cls) -> str:
        """インストール手順を取得"""
        return f"""Ghostscriptがインストールされていません。

【インストール手順】
1. 以下のURLからGhostscriptをダウンロード:
   {cls.DOWNLOAD_URL}

2. 「AGPL Release」から最新版をダウンロード
   - 64bit Windows: Ghostscript X.XX.X for Windows (64 bit)
   - 32bit Windows: Ghostscript X.XX.X for Windows (32 bit)

3. ダウンロードしたインストーラーを実行

4. インストール完了後、このアプリを再起動するか
   設定タブで「自動検出」ボタンを押してください"""


def auto_configure_ghostscript(config) -> bool:
    """
    Ghostscriptを自動検出して設定に登録

    Args:
        config: ConfigLoaderインスタンス

    Returns:
        bool: 設定に成功した場合True
    """
    gs_path = GhostscriptManager.find_ghostscript()

    if gs_path and GhostscriptManager.verify_ghostscript(gs_path):
        config.set('ghostscript', 'executable', value=gs_path)
        config.save_config()
        logger.info(f"Ghostscriptパスを設定しました: {gs_path}")
        return True

    return False
