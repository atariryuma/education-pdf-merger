"""
インストール後の初期設定スクリプト

Ghostscriptの検出と設定ファイルの更新を行う
"""
import os
import sys
import json
import subprocess
import winreg
from typing import Optional


def find_ghostscript() -> Optional[str]:
    """Ghostscriptを検索"""
    # 1. レジストリから検索
    gs_path = find_from_registry()
    if gs_path:
        return gs_path

    # 2. 既知のパスから検索
    gs_path = find_from_known_paths()
    if gs_path:
        return gs_path

    return None


def find_from_registry() -> Optional[str]:
    """レジストリからGhostscriptを検索"""
    registry_paths = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\GPL Ghostscript"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\GPL Ghostscript"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\GPL Ghostscript"),
    ]
    exe_names = ["gswin64c.exe", "gswin32c.exe"]

    for hkey, subkey in registry_paths:
        try:
            with winreg.OpenKey(hkey, subkey) as key:
                versions = []
                i = 0
                while True:
                    try:
                        version = winreg.EnumKey(key, i)
                        versions.append(version)
                        i += 1
                    except OSError:
                        break

                if versions:
                    versions.sort(reverse=True)
                    latest_version = versions[0]
                    with winreg.OpenKey(key, latest_version) as version_key:
                        gs_dir, _ = winreg.QueryValueEx(version_key, "GS_DLL")
                        bin_dir = os.path.dirname(gs_dir)
                        for exe_name in exe_names:
                            exe_path = os.path.join(bin_dir, exe_name)
                            if os.path.exists(exe_path):
                                return exe_path
        except OSError:
            continue

    return None


def find_from_known_paths() -> Optional[str]:
    """既知のパスからGhostscriptを検索"""
    search_paths = [
        r"C:\Program Files\gs",
        r"C:\Program Files (x86)\gs",
        r"C:\gs",
    ]
    exe_names = ["gswin64c.exe", "gswin32c.exe"]

    for base_path in search_paths:
        if not os.path.exists(base_path):
            continue

        try:
            subdirs = [d for d in os.listdir(base_path)
                      if os.path.isdir(os.path.join(base_path, d)) and d.startswith("gs")]
            subdirs.sort(reverse=True)

            for subdir in subdirs:
                bin_dir = os.path.join(base_path, subdir, "bin")
                if os.path.exists(bin_dir):
                    for exe_name in exe_names:
                        exe_path = os.path.join(bin_dir, exe_name)
                        if os.path.exists(exe_path):
                            return exe_path
        except OSError:
            continue

    return None


def verify_ghostscript(gs_path: str) -> bool:
    """Ghostscriptの動作確認"""
    if not gs_path or not os.path.exists(gs_path):
        return False

    try:
        result = subprocess.run(
            [gs_path, "--version"],
            capture_output=True,
            text=True,
            timeout=10
        )
        return result.returncode == 0
    except Exception:
        return False


def update_config(app_dir: str, gs_path: str) -> bool:
    """設定ファイルを更新"""
    config_path = os.path.join(app_dir, "config.json")

    if not os.path.exists(config_path):
        return False

    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)

        # Ghostscriptパスを設定
        if 'ghostscript' not in config:
            config['ghostscript'] = {}
        config['ghostscript']['executable'] = gs_path

        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)

        return True
    except Exception:
        return False


def main():
    """メイン処理"""
    if len(sys.argv) < 2:
        print("Usage: post_install.py <app_dir>")
        sys.exit(1)

    app_dir = sys.argv[1]

    # Ghostscriptを検索
    gs_path = find_ghostscript()

    if gs_path and verify_ghostscript(gs_path):
        # 設定ファイルを更新
        if update_config(app_dir, gs_path):
            print(f"OK:{gs_path}")
            sys.exit(0)
        else:
            print("ERROR:config_update_failed")
            sys.exit(2)
    else:
        print("NOT_FOUND")
        sys.exit(1)


if __name__ == "__main__":
    main()
