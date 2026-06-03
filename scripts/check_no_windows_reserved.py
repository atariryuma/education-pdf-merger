"""
Windows予約名ファイルを検出するpre-commitチェックスクリプト

Windowsには NUL, CON, PRN, AUX, COM1-9, LPT1-9 等の予約デバイス名がある。
これらと同名のファイルが作成されると、Windowsツールチェーン（mypy等）が
ファイルシステム走査時に予期しないエラーを起こす。

このスクリプトはリポジトリルートから走査し、予約名ファイルを検出して
exit code 1 で失敗させる。
"""
from __future__ import annotations

import sys
from pathlib import Path

# Windows予約名（ベース名のみ。拡張子付きでも予約）
RESERVED_NAMES = {
    "nul", "con", "prn", "aux",
    *(f"com{i}" for i in range(1, 10)),
    *(f"lpt{i}" for i in range(1, 10)),
}

# 走査対象から除外するディレクトリ
EXCLUDE_DIRS = {".git", "venv", ".venv", "build", "dist", "__pycache__", "node_modules"}


def _is_reserved(name: str) -> bool:
    """ファイル名がWindows予約名かを判定（拡張子を除いたstemで比較）"""
    stem = name.split(".", 1)[0].lower()
    return stem in RESERVED_NAMES


def find_reserved_files(root: Path) -> list[Path]:
    """ルート以下から予約名ファイルを検索"""
    offenders: list[Path] = []
    for path in root.rglob("*"):
        if any(part in EXCLUDE_DIRS for part in path.parts):
            continue
        if path.is_file() and _is_reserved(path.name):
            offenders.append(path)
    return offenders


def main() -> int:
    root = Path(__file__).resolve().parent.parent
    offenders = find_reserved_files(root)
    if not offenders:
        return 0

    print("❌ Windows予約名ファイルを検出しました:", file=sys.stderr)
    for path in offenders:
        print(f"  - {path.relative_to(root)}", file=sys.stderr)
    print(
        "\nこれらのファイルはWindowsで誤動作の原因となります。削除してください。\n"
        "例: del nul （PowerShellでは Remove-Item -LiteralPath nul -Force）",
        file=sys.stderr,
    )
    return 1


if __name__ == "__main__":
    sys.exit(main())
