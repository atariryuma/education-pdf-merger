"""
設定ファイル検証ユーティリティ

config.jsonの完全性と妥当性を検証します。
"""
import logging
from pathlib import Path
from typing import List, Tuple, Optional
from dataclasses import dataclass
from enum import Enum

from config_loader import ConfigLoader
from ghostscript_detector import GhostscriptDetector

logger = logging.getLogger(__name__)


class ValidationLevel(Enum):
    """検証レベル"""
    ERROR = "error"      # 必須項目の欠如（アプリが動作しない）
    WARNING = "warning"  # 推奨項目の欠如（一部機能が使えない）
    INFO = "info"        # 情報（最適化の提案など）


@dataclass
class ValidationResult:
    """検証結果"""
    level: ValidationLevel
    message: str
    field: Optional[str] = None  # 問題のあるフィールド名


class ConfigValidator:
    """設定ファイル検証クラス

    ベストプラクティス:
    - 必須項目のチェック
    - パスの存在確認
    - 値の妥当性検証
    - 推奨設定の確認
    """

    def __init__(self, config: ConfigLoader) -> None:
        """
        Args:
            config: 検証する設定オブジェクト
        """
        self.config = config
        self.results: List[ValidationResult] = []

    def validate_all(self) -> Tuple[bool, List[ValidationResult]]:
        """全ての設定を検証

        Returns:
            (is_valid, results): 有効性フラグと検証結果のリスト
        """
        logger.info("設定ファイルの検証を開始")
        self.results = []

        # 必須項目の検証
        self._validate_required_fields()

        # パスの検証
        self._validate_paths()

        # Ghostscriptの検証
        self._validate_ghostscript()

        # Excelファイルの検証
        self._validate_excel_files()

        # エラーレベルの結果があれば無効
        has_errors = any(r.level == ValidationLevel.ERROR for r in self.results)
        is_valid = not has_errors

        if is_valid:
            logger.info("設定ファイルの検証が完了しました（エラーなし）")
        else:
            logger.warning(f"設定ファイルの検証が完了しました（エラー数: {sum(1 for r in self.results if r.level == ValidationLevel.ERROR)}）")

        return is_valid, self.results

    def _validate_required_fields(self) -> None:
        """必須フィールドの検証"""
        # 年度（西暦のみ）
        year = self.config.year
        if not year or year.strip() == "":
            self.results.append(ValidationResult(
                level=ValidationLevel.ERROR,
                message="年度が設定されていません（例: 2026）",
                field="year"
            ))

        # year_shortは自動計算されるため、検証は不要（INFO）
        year_short = self.config.year_short
        if not year_short or year_short.strip() == "":
            self.results.append(ValidationResult(
                level=ValidationLevel.INFO,
                message="年度（短縮形）は西暦から自動計算されます",
                field="year_short"
            ))

        # Google Driveパス
        gdrive = self.config.get('base_paths', 'google_drive')
        if not gdrive or gdrive.strip() == "":
            self.results.append(ValidationResult(
                level=ValidationLevel.ERROR,
                message="作業フォルダ（Google Drive等）が設定されていません",
                field="base_paths.google_drive"
            ))

    def _validate_paths(self) -> None:
        """パスの妥当性検証"""
        # Google Driveパス
        gdrive = self.config.get('base_paths', 'google_drive')
        if gdrive and gdrive.strip():
            gdrive_path = Path(gdrive)
            if not gdrive_path.exists():
                self.results.append(ValidationResult(
                    level=ValidationLevel.WARNING,
                    message=f"作業フォルダが見つかりません: {gdrive}",
                    field="base_paths.google_drive"
                ))
            elif not gdrive_path.is_dir():
                self.results.append(ValidationResult(
                    level=ValidationLevel.ERROR,
                    message=f"作業フォルダがディレクトリではありません: {gdrive}",
                    field="base_paths.google_drive"
                ))

        # 一時フォルダ（自動生成されるのでWARNING）
        temp = self.config.get('base_paths', 'local_temp')
        if temp and temp.strip():
            temp_path = Path(temp)
            if not temp_path.exists():
                self.results.append(ValidationResult(
                    level=ValidationLevel.INFO,
                    message=f"一時フォルダが存在しません（初回実行時に自動作成されます）: {temp}",
                    field="base_paths.local_temp"
                ))

        # フォントファイル
        font = self.config.get('fonts', 'mincho')
        if font:
            font_path = Path(font)
            if not font_path.exists():
                self.results.append(ValidationResult(
                    level=ValidationLevel.WARNING,
                    message=f"フォントファイルが見つかりません: {font}",
                    field="fonts.mincho"
                ))

    def _validate_ghostscript(self) -> None:
        """Ghostscript設定の検証"""
        gs_path = self.config.get('ghostscript', 'executable')

        if not gs_path or gs_path.strip() == "":
            # 自動検出を試みる
            detected = GhostscriptDetector.detect()
            if detected:
                self.results.append(ValidationResult(
                    level=ValidationLevel.INFO,
                    message=f"Ghostscriptが自動検出されました: {detected}",
                    field="ghostscript.executable"
                ))
            else:
                self.results.append(ValidationResult(
                    level=ValidationLevel.WARNING,
                    message="Ghostscriptが設定されていません（PDF圧縮機能が使用できません）",
                    field="ghostscript.executable"
                ))
        else:
            # パスの妥当性を検証
            if not GhostscriptDetector.validate_ghostscript(gs_path):
                self.results.append(ValidationResult(
                    level=ValidationLevel.WARNING,
                    message=f"Ghostscriptパスが無効です: {gs_path}",
                    field="ghostscript.executable"
                ))

    def _validate_excel_files(self) -> None:
        """Excelファイル設定の検証"""
        # Excel file paths are now session-based (not stored in config.json)
        # Only sheet names are stored in config, which don't require validation here
        pass

    def get_missing_required_fields(self) -> List[str]:
        """必須項目の欠如リストを取得

        Returns:
            欠如している必須フィールド名のリスト
        """
        return [
            r.field for r in self.results
            if r.level == ValidationLevel.ERROR and r.field
        ]

    def has_errors(self) -> bool:
        """エラーレベルの問題があるか確認

        Returns:
            エラーがある場合True
        """
        return any(r.level == ValidationLevel.ERROR for r in self.results)

    def has_warnings(self) -> bool:
        """警告レベルの問題があるか確認

        Returns:
            警告がある場合True
        """
        return any(r.level == ValidationLevel.WARNING for r in self.results)

    def get_summary(self) -> str:
        """検証結果のサマリーを取得

        Returns:
            検証結果のテキストサマリー
        """
        errors = [r for r in self.results if r.level == ValidationLevel.ERROR]
        warnings = [r for r in self.results if r.level == ValidationLevel.WARNING]
        infos = [r for r in self.results if r.level == ValidationLevel.INFO]

        summary_lines = []

        if errors:
            summary_lines.append(f"[ERROR] エラー ({len(errors)}件):")
            for r in errors:
                summary_lines.append(f"  - {r.message}")

        if warnings:
            summary_lines.append(f"\n[WARNING] 警告 ({len(warnings)}件):")
            for r in warnings:
                summary_lines.append(f"  - {r.message}")

        if infos:
            summary_lines.append(f"\n[INFO] 情報 ({len(infos)}件):")
            for r in infos:
                summary_lines.append(f"  - {r.message}")

        if not summary_lines:
            return "[OK] 設定に問題はありません"

        return "\n".join(summary_lines)
