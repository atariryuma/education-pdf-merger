"""
Ghostscriptユーティリティモジュール

Ghostscriptの自動検出、パス検証、動作確認、インストール案内を提供
"""
import logging
import os
import subprocess

try:
    import winreg
except ImportError:
    winreg = None  # type: ignore[assignment]
from pathlib import Path
from typing import Callable, Optional, List, Tuple

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
        "gs.exe",  # 汎用名
    ]

    # 標準インストールパス
    STANDARD_PATHS = [
        r"C:\Program Files\gs",
        r"C:\Program Files (x86)\gs",
    ]

    @classmethod
    def _get_registry_keys(cls) -> list:
        """レジストリキーのリストを返す（winreg利用可能時のみ）"""
        if winreg is None:
            return []
        return [
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
        gs_dll = os.environ.get("GS_DLL")
        if gs_dll and Path(gs_dll).exists():
            dll_dir = Path(gs_dll).parent
            for exe_name in cls.GS_EXECUTABLES:
                exe_path = dll_dir / exe_name
                if exe_path.exists():
                    return str(exe_path)

        # GS_LIB環境変数（セミコロン区切りの複数パスに対応）
        gs_lib = os.environ.get("GS_LIB")
        if gs_lib:
            for lib_path_str in gs_lib.split(";"):
                lib_path_str = lib_path_str.strip()
                if not lib_path_str:
                    continue
                lib_dir = Path(lib_path_str)
                if lib_dir.exists():
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
        if winreg is None:
            logger.debug("winregが利用できないため、レジストリ検索をスキップ")
            return None

        found_versions: List[Tuple[str, str]] = []

        for root, key_path in cls._get_registry_keys():
            try:
                with winreg.OpenKey(root, key_path) as key:
                    i = 0
                    while True:
                        try:
                            version = winreg.EnumKey(key, i)
                            try:
                                with winreg.OpenKey(
                                    root, f"{key_path}\\{version}"
                                ) as ver_key:
                                    try:
                                        gs_dll = winreg.QueryValueEx(ver_key, "GS_DLL")[
                                            0
                                        ]
                                        if Path(gs_dll).exists():
                                            dll_dir = Path(gs_dll).parent
                                            for exe_name in cls.GS_EXECUTABLES:
                                                exe_path = dll_dir / exe_name
                                                if exe_path.exists():
                                                    found_versions.append(
                                                        (version, str(exe_path))
                                                    )
                                                    break
                                    except FileNotFoundError:
                                        pass

                                    try:
                                        gs_lib = winreg.QueryValueEx(ver_key, "GS_LIB")[
                                            0
                                        ]
                                        lib_dir = Path(gs_lib)
                                        if lib_dir.exists():
                                            bin_dir = lib_dir.parent / "bin"
                                            if bin_dir.exists():
                                                for exe_name in cls.GS_EXECUTABLES:
                                                    exe_path = bin_dir / exe_name
                                                    if exe_path.exists():
                                                        found_versions.append(
                                                            (version, str(exe_path))
                                                        )
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

        if found_versions:
            found_versions.sort(reverse=True, key=lambda x: cls._parse_version(x[0]))
            logger.debug(
                f"レジストリから検出されたGhostscriptバージョン: {[v[0] for v in found_versions]}"
            )
            return found_versions[0][1]

        return None

    @classmethod
    def _parse_version(cls, version_str: str) -> Tuple[int, ...]:
        """バージョン文字列を解析してタプルに変換"""
        try:
            parts = version_str.split(".")
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

            for exe_name in cls.GS_EXECUTABLES:
                for exe_path in base.rglob(exe_name):
                    if exe_path.parent.name == "bin":
                        return str(exe_path)

        return None

    @classmethod
    def _check_path_env(cls) -> Optional[str]:
        """PATH環境変数から検出"""
        path_env = os.environ.get("PATH", "")
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

        if path.name.lower() not in [exe.lower() for exe in cls.GS_EXECUTABLES]:
            logger.warning(f"Ghostscript実行ファイル名が不正です: {path.name}")
            return False

        logger.debug(f"Ghostscriptパス検証OK: {gs_path}")
        return True

    @classmethod
    def verify(cls, gs_path: str) -> bool:
        """
        Ghostscriptが正常に動作するか確認（subprocessで--versionを実行）

        パス検証を事前に実施し、信頼できるパスのみ実行する。

        Args:
            gs_path: Ghostscript実行ファイルのパス

        Returns:
            bool: 正常に動作する場合True
        """
        if not gs_path or not os.path.exists(gs_path):
            return False

        # パスの妥当性を検証（ファイル名が既知のGS実行ファイルであること）
        if not cls.validate_ghostscript(gs_path):
            return False

        try:
            result = subprocess.run(
                [gs_path, "--version"], capture_output=True, text=True, timeout=10
            )
            if result.returncode == 0:
                version = result.stdout.strip()
                logger.info(f"Ghostscript バージョン: {version}")
                return True
        except subprocess.TimeoutExpired:
            logger.warning(f"Ghostscript検証タイムアウト: {gs_path}")
        except FileNotFoundError:
            logger.warning(f"Ghostscriptが見つかりません: {gs_path}")
        except OSError as e:
            logger.warning(f"Ghostscript検証エラー: {e}")

        return False

    @classmethod
    def get_install_instructions(cls) -> str:
        """インストール手順を取得"""
        return (
            "Ghostscriptが見つかりませんでした。\n\n"
            "設定タブの「⬇ インストール」ボタンで自動インストールできます。"
        )


class GhostscriptInstaller:
    """Ghostscriptの自動ダウンロード・インストール"""

    # GitHub APIからリリース情報を取得
    GITHUB_API_URL = "https://api.github.com/repos/ArtifexSoftware/ghostpdl-downloads/releases/latest"

    # フォールバック: 既知の安定版URL
    FALLBACK_URL_64 = "https://github.com/ArtifexSoftware/ghostpdl-downloads/releases/download/gs10040/gs10040w64.exe"
    FALLBACK_URL_32 = "https://github.com/ArtifexSoftware/ghostpdl-downloads/releases/download/gs10040/gs10040w32.exe"

    @classmethod
    def download_and_install(
        cls, progress_callback: Optional[Callable] = None
    ) -> tuple:
        """
        Ghostscriptをダウンロードしてサイレントインストール

        Args:
            progress_callback: 進捗メッセージ用コールバック

        Returns:
            (success: bool, message: str, gs_path: Optional[str])
        """
        import platform
        import tempfile
        import urllib.request
        import urllib.error

        is_64bit = platform.machine().endswith("64")

        # ダウンロードURLを決定
        download_url = cls._get_download_url(is_64bit)
        if not download_url:
            download_url = cls.FALLBACK_URL_64 if is_64bit else cls.FALLBACK_URL_32

        if progress_callback:
            progress_callback("⬇ ダウンロード中...")

        # 一時ファイルにダウンロード
        try:
            installer_path = os.path.join(tempfile.gettempdir(), "gs_installer.exe")
            urllib.request.urlretrieve(download_url, installer_path)
            logger.info(f"Ghostscriptインストーラーをダウンロード: {installer_path}")
        except (urllib.error.URLError, OSError) as e:
            logger.error(f"ダウンロード失敗: {e}")
            return (
                False,
                f"ダウンロードに失敗しました。\n\nインターネット接続を確認してください。\n詳細: {e}",
                None,
            )

        if progress_callback:
            progress_callback("📦 インストール中...")

        # サイレントインストール実行
        try:
            result = subprocess.run(
                [installer_path, "/S"], capture_output=True, text=True, timeout=120
            )
            if result.returncode != 0:
                logger.error(f"インストーラー終了コード: {result.returncode}")
                return (
                    False,
                    f"インストールに失敗しました（終了コード: {result.returncode}）",
                    None,
                )
        except subprocess.TimeoutExpired:
            return False, "インストールがタイムアウトしました（120秒）", None
        except OSError as e:
            return False, f"インストーラーの実行に失敗しました: {e}", None
        finally:
            # インストーラーを削除
            try:
                os.remove(installer_path)
            except OSError:
                pass

        if progress_callback:
            progress_callback("🔍 検出中...")

        # インストール後に自動検出
        gs_path = GhostscriptDetector.detect()
        if gs_path and GhostscriptDetector.verify(gs_path):
            return (
                True,
                f"Ghostscriptのインストールが完了しました。\n\nパス: {gs_path}",
                gs_path,
            )
        else:
            return (
                False,
                "インストールは完了しましたが、Ghostscriptが検出できませんでした。\n再起動後に自動検出を試してください。",
                None,
            )

    @classmethod
    def _get_download_url(cls, is_64bit: bool) -> Optional[str]:
        """GitHub APIから最新リリースのダウンロードURLを取得"""
        import urllib.request
        import urllib.error
        import json

        try:
            req = urllib.request.Request(
                cls.GITHUB_API_URL, headers={"User-Agent": "PDFMergeSystem"}
            )
            with urllib.request.urlopen(req, timeout=10) as response:
                data = json.loads(response.read().decode())

            suffix = "w64.exe" if is_64bit else "w32.exe"
            for asset in data.get("assets", []):
                name = asset.get("name", "")
                if name.endswith(suffix) and "gs" in name.lower():
                    url: str = asset.get("browser_download_url", "")
                    if url:
                        logger.info(f"最新Ghostscript URL: {url}")
                        return url
        except (urllib.error.URLError, json.JSONDecodeError, OSError) as e:
            logger.warning(f"GitHub APIからURL取得失敗（フォールバック使用）: {e}")

        return None
