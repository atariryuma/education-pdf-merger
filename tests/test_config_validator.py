"""
ConfigValidator のユニットテスト

設定ファイルの検証ロジックをテスト
"""
import pytest
from unittest.mock import MagicMock, patch

from infrastructure.config_validator import (
    ConfigValidator, ValidationLevel, ValidationResult
)


def _make_config(**overrides):
    """テスト用のConfigLoaderモックを作成"""
    config = MagicMock()
    config.year = overrides.get("year", "2026")
    config.year_short = overrides.get("year_short", "R8")

    def get_side_effect(section, key=None):
        defaults = {
            ('base_paths', 'google_drive'): overrides.get("gdrive", "C:\\Users\\test\\GoogleDrive"),
            ('base_paths', 'local_temp'): overrides.get("local_temp", "C:\\temp"),
            ('fonts', 'mincho'): overrides.get("font", None),
            ('ghostscript', 'executable'): overrides.get("gs", ""),
        }
        return defaults.get((section, key), "")

    config.get = MagicMock(side_effect=get_side_effect)
    return config


@pytest.mark.unit
class TestValidationResult:
    """ValidationResult データクラスのテスト"""

    def test_create_error(self):
        r = ValidationResult(ValidationLevel.ERROR, "エラーメッセージ", "field")
        assert r.level == ValidationLevel.ERROR
        assert r.message == "エラーメッセージ"
        assert r.field == "field"

    def test_create_without_field(self):
        r = ValidationResult(ValidationLevel.WARNING, "警告")
        assert r.field is None


@pytest.mark.unit
class TestValidateRequiredFields:
    """必須フィールド検証のテスト"""

    def test_valid_config_no_errors(self):
        """有効な設定ではエラーなし"""
        config = _make_config()
        validator = ConfigValidator(config)
        validator._validate_required_fields()
        errors = [r for r in validator.results if r.level == ValidationLevel.ERROR]
        assert len(errors) == 0

    def test_missing_year(self):
        """年度が空"""
        config = _make_config(year="")
        validator = ConfigValidator(config)
        validator._validate_required_fields()
        errors = [r for r in validator.results if r.level == ValidationLevel.ERROR]
        assert any("年度" in r.message for r in errors)

    def test_none_year(self):
        """年度がNone"""
        config = _make_config(year=None)
        validator = ConfigValidator(config)
        validator._validate_required_fields()
        errors = [r for r in validator.results if r.level == ValidationLevel.ERROR]
        assert any(r.field == "year" for r in errors)

    def test_missing_gdrive(self):
        """Google Driveパスが空"""
        config = _make_config(gdrive="")
        validator = ConfigValidator(config)
        validator._validate_required_fields()
        errors = [r for r in validator.results if r.level == ValidationLevel.ERROR]
        assert any("作業フォルダ" in r.message for r in errors)

    def test_missing_year_short_is_info(self):
        """year_short未設定はINFOレベル"""
        config = _make_config(year_short="")
        validator = ConfigValidator(config)
        validator._validate_required_fields()
        infos = [r for r in validator.results if r.level == ValidationLevel.INFO]
        assert any("自動計算" in r.message for r in infos)


@pytest.mark.unit
class TestValidatePaths:
    """パス検証のテスト"""

    def test_existing_gdrive_path(self, tmp_path):
        """存在するGoogle Driveパス"""
        config = _make_config(gdrive=str(tmp_path))
        validator = ConfigValidator(config)
        validator._validate_paths()
        warnings = [r for r in validator.results if r.level == ValidationLevel.WARNING]
        assert len(warnings) == 0

    def test_nonexistent_gdrive_path(self):
        """存在しないGoogle Driveパス"""
        config = _make_config(gdrive=r"Z:\nonexistent\path")
        validator = ConfigValidator(config)
        validator._validate_paths()
        warnings = [r for r in validator.results if r.level == ValidationLevel.WARNING]
        assert any("見つかりません" in r.message for r in warnings)

    def test_gdrive_path_is_file(self, tmp_path):
        """Google Driveパスがファイル"""
        file = tmp_path / "not_a_dir.txt"
        file.write_text("test")
        config = _make_config(gdrive=str(file))
        validator = ConfigValidator(config)
        validator._validate_paths()
        errors = [r for r in validator.results if r.level == ValidationLevel.ERROR]
        assert any("ディレクトリではありません" in r.message for r in errors)

    def test_nonexistent_temp_is_info(self):
        """存在しない一時フォルダはINFO"""
        config = _make_config(local_temp=r"Z:\nonexistent\temp")
        validator = ConfigValidator(config)
        validator._validate_paths()
        infos = [r for r in validator.results if r.level == ValidationLevel.INFO]
        assert any("自動作成" in r.message for r in infos)


@pytest.mark.unit
class TestValidateGhostscript:
    """Ghostscript検証のテスト"""

    @patch("infrastructure.config_validator.GhostscriptDetector")
    def test_gs_not_set_detected(self, mock_gs):
        """GS未設定で自動検出成功"""
        mock_gs.detect.return_value = r"C:\gs\bin\gswin64c.exe"
        config = _make_config(gs="")
        validator = ConfigValidator(config)
        validator._validate_ghostscript()
        infos = [r for r in validator.results if r.level == ValidationLevel.INFO]
        assert any("自動検出" in r.message for r in infos)

    @patch("infrastructure.config_validator.GhostscriptDetector")
    def test_gs_not_set_not_detected(self, mock_gs):
        """GS未設定で自動検出失敗"""
        mock_gs.detect.return_value = None
        config = _make_config(gs="")
        validator = ConfigValidator(config)
        validator._validate_ghostscript()
        warnings = [r for r in validator.results if r.level == ValidationLevel.WARNING]
        assert any("設定されていません" in r.message for r in warnings)

    @patch("infrastructure.config_validator.GhostscriptDetector")
    def test_gs_invalid_path(self, mock_gs):
        """GS設定済みだがパスが無効"""
        mock_gs.validate_ghostscript.return_value = False
        config = _make_config(gs=r"C:\invalid\gs.exe")
        validator = ConfigValidator(config)
        validator._validate_ghostscript()
        warnings = [r for r in validator.results if r.level == ValidationLevel.WARNING]
        assert any("無効です" in r.message for r in warnings)


@pytest.mark.unit
class TestValidateAll:
    """validate_all 統合テスト"""

    @patch("infrastructure.config_validator.GhostscriptDetector")
    def test_valid_config(self, mock_gs, tmp_path):
        """全て有効な設定"""
        mock_gs.detect.return_value = None
        config = _make_config(gdrive=str(tmp_path))
        validator = ConfigValidator(config)
        is_valid, results = validator.validate_all()
        assert is_valid is True
        assert not validator.has_errors()

    @patch("infrastructure.config_validator.GhostscriptDetector")
    def test_invalid_config(self, mock_gs):
        """エラーのある設定"""
        mock_gs.detect.return_value = None
        config = _make_config(year="", gdrive="")
        validator = ConfigValidator(config)
        is_valid, results = validator.validate_all()
        assert is_valid is False
        assert validator.has_errors()
        assert len(validator.get_missing_required_fields()) >= 2


@pytest.mark.unit
class TestGetSummary:
    """get_summary のテスト"""

    def test_no_issues(self):
        config = _make_config()
        validator = ConfigValidator(config)
        validator.results = []
        assert "問題はありません" in validator.get_summary()

    def test_with_errors(self):
        config = _make_config()
        validator = ConfigValidator(config)
        validator.results = [
            ValidationResult(ValidationLevel.ERROR, "エラー1", "field1")
        ]
        summary = validator.get_summary()
        assert "[ERROR]" in summary
        assert "エラー1" in summary

    def test_with_mixed(self):
        config = _make_config()
        validator = ConfigValidator(config)
        validator.results = [
            ValidationResult(ValidationLevel.ERROR, "err"),
            ValidationResult(ValidationLevel.WARNING, "warn"),
            ValidationResult(ValidationLevel.INFO, "info"),
        ]
        summary = validator.get_summary()
        assert "[ERROR]" in summary
        assert "[WARNING]" in summary
        assert "[INFO]" in summary
