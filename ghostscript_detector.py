"""
Ghostscript自動検出ユーティリティ

Windowsシステム上のGhostscriptインストールを自動検出します。
"""
import logging
import os
import winreg
from pathlib import Path
from typing import Optional, List, Tuple

logger = logging.getLogger(__name__)


class GhostscriptDetector:
    """Ghostscript自動検出クラス

    ベストプラクティス:
    - Windowsレジストリからの検出
    - 標準インストールパスの検索
    - 環境変数の確認
    - PATH環境変数の検索

    参考: https://www.biopdf.com/guide/detecting_ghostscript.php
    """

    # Ghostscript実行ファイル名（優先度順）
    GS_EXECUTABLES = [
        "gswin64c.exe",  # 64bit コンソール版（推奨）
        "gswin32c.exe",  # 32bit コンソール版
        "gs.exe",        # 汎用名
    ]

    # 標準インストールパス
    STANDARD_PATHS = [
        r"C:\Program Files\gs",
        r"C:\Program Files (x86)\gs",
    ]

    # レジストリキー
    REGISTRY_KEYS = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\GPL Ghostscript"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\GPL Ghostscript"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\AFPL Ghostscript"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\AFPL Ghostscript"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\GPL Ghostscript"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\AFPL Ghostscript"),
    ]

    @classmethod
    def detect(cls) -> Optional[str]:
        """Ghostscriptパスを自動検出

        Returns:
            検出されたGhostscript実行ファイルのパス、見つからない場合はNone

        Note:
            以下の順序で検索します:
            1. 環境変数 (GS_DLL, GS_LIB)
            2. Windowsレジストリ
            3. 標準インストールパス
            4. PATH環境変数
        """
        logger.info("Ghostscriptの自動検出を開始")

        # 1. 環境変数チェック
        gs_path = cls._check_environment_variables()
        if gs_path:
            logger.info(f"環境変数からGhostscriptを検出: {gs_path}")
            return gs_path

        # 2. レジストリチェック
        gs_path = cls._check_registry()
        if gs_path:
            logger.info(f"レジストリからGhostscriptを検出: {gs_path}")
            return gs_path

        # 3. 標準パス検索
        gs_path = cls._check_standard_paths()
        if gs_path:
            logger.info(f"標準パスからGhostscriptを検出: {gs_path}")
            return gs_path

        # 4. PATH環境変数検索
        gs_path = cls._check_path_env()
        if gs_path:
            logger.info(f"PATH環境変数からGhostscriptを検出: {gs_path}")
            return gs_path

        logger.warning("Ghostscriptが見つかりませんでした")
        return None

    @classmethod
    def _check_environment_variables(cls) -> Optional[str]:
        """環境変数からGhostscriptを検出"""
        # GS_DLL環境変数
        gs_dll = os.environ.get('GS_DLL')
        if gs_dll and Path(gs_dll).exists():
            # DLLパスから実行ファイルパスを推定
            dll_dir = Path(gs_dll).parent
            for exe_name in cls.GS_EXECUTABLES:
                exe_path = dll_dir / exe_name
                if exe_path.exists():
                    return str(exe_path)

        # GS_LIB環境変数
        gs_lib = os.environ.get('GS_LIB')
        if gs_lib:
            # LIBパスから実行ファイルを検索
            lib_dir = Path(gs_lib)
            if lib_dir.exists():
                # 親ディレクトリのbinフォルダを探す
                bin_dir = lib_dir.parent / "bin"
                if bin_dir.exists():
                    for exe_name in cls.GS_EXECUTABLES:
                        exe_path = bin_dir / exe_name
                        if exe_path.exists():
                            return str(exe_path)

        return None

    @classmethod
    def _check_registry(cls) -> Optional[str]:
        """Windowsレジストリから検出"""
        found_versions: List[Tuple[str, str]] = []  # (version, path)

        for root, key_path in cls.REGISTRY_KEYS:
            try:
                with winreg.OpenKey(root, key_path) as key:
                    # バージョン一覧を取得
                    i = 0
                    while True:
                        try:
                            version = winreg.EnumKey(key, i)
                            # バージョンキーからGS_DLLまたはGS_LIBを取得
                            try:
                                with winreg.OpenKey(root, f"{key_path}\\{version}") as ver_key:
                                    # GS_DLL
                                    try:
                                        gs_dll = winreg.QueryValueEx(ver_key, "GS_DLL")[0]
                                        if Path(gs_dll).exists():
                                            dll_dir = Path(gs_dll).parent
                                            for exe_name in cls.GS_EXECUTABLES:
                                                exe_path = dll_dir / exe_name
                                                if exe_path.exists():
                                                    found_versions.append((version, str(exe_path)))
                                                    break
                                    except FileNotFoundError:
                                        pass

                                    # GS_LIB
                                    try:
                                        gs_lib = winreg.QueryValueEx(ver_key, "GS_LIB")[0]
                                        lib_dir = Path(gs_lib)
                                        if lib_dir.exists():
                                            bin_dir = lib_dir.parent / "bin"
                                            if bin_dir.exists():
                                                for exe_name in cls.GS_EXECUTABLES:
                                                    exe_path = bin_dir / exe_name
                                                    if exe_path.exists():
                                                        found_versions.append((version, str(exe_path)))
                                                        break
                                    except FileNotFoundError:
                                        pass
                            except OSError:
                                pass
                            i += 1
                        except OSError:
                            break
            except FileNotFoundError:
                continue
            except OSError as e:
                logger.debug(f"レジストリキー {root}\\{key_path} の読み取りに失敗: {e}")
                continue

        # 最新バージョンを返す
        if found_versions:
            # バージョン番号でソート（降順）
            found_versions.sort(reverse=True, key=lambda x: cls._parse_version(x[0]))
            logger.debug(f"レジストリから検出されたGhostscriptバージョン: {[v[0] for v in found_versions]}")
            return found_versions[0][1]

        return None

    @classmethod
    def _parse_version(cls, version_str: str) -> Tuple[int, ...]:
        """バージョン文字列を解析してタプルに変換

        Args:
            version_str: バージョン文字列（例: "10.05.0", "9.54"）

        Returns:
            バージョンタプル（例: (10, 5, 0), (9, 54, 0)）
        """
        try:
            parts = version_str.split('.')
            return tuple(int(p) for p in parts)
        except (ValueError, AttributeError):
            return (0,)

    @classmethod
    def _check_standard_paths(cls) -> Optional[str]:
        """標準インストールパスから検出"""
        for base_path in cls.STANDARD_PATHS:
            base = Path(base_path)
            if not base.exists():
                continue

            # gsXX.XX.XX/bin/gswinXXc.exe のパターンを探す
            for exe_name in cls.GS_EXECUTABLES:
                for exe_path in base.rglob(exe_name):
                    # binディレクトリ内にあることを確認
                    if exe_path.parent.name == "bin":
                        return str(exe_path)

        return None

    @classmethod
    def _check_path_env(cls) -> Optional[str]:
        """PATH環境変数から検出"""
        path_env = os.environ.get('PATH', '')
        for path_dir in path_env.split(os.pathsep):
            if not path_dir:
                continue

            for exe_name in cls.GS_EXECUTABLES:
                exe_path = Path(path_dir) / exe_name
                if exe_path.exists():
                    return str(exe_path)

        return None

    @classmethod
    def validate_ghostscript(cls, gs_path: str) -> bool:
        """Ghostscriptパスの妥当性を検証

        Args:
            gs_path: Ghostscript実行ファイルのパス

        Returns:
            パスが有効な場合True
        """
        if not gs_path:
            return False

        path = Path(gs_path)
        if not path.exists():
            logger.warning(f"Ghostscriptパスが存在しません: {gs_path}")
            return False

        if not path.is_file():
            logger.warning(f"Ghostscriptパスがファイルではありません: {gs_path}")
            return False

        # 実行ファイル名の確認
        if path.name.lower() not in [exe.lower() for exe in cls.GS_EXECUTABLES]:
            logger.warning(f"Ghostscript実行ファイル名が不正です: {path.name}")
            return False

        logger.info(f"Ghostscriptパスが有効です: {gs_path}")
        return True
