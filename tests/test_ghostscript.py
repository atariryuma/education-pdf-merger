"""
GhostscriptDetector のユニットテスト

自動検出、パス検証、動作確認をテスト
"""
import os
import pytest
from unittest.mock import patch, MagicMock


from infrastructure.ghostscript import GhostscriptDetector


@pytest.mark.unit
class TestValidateGhostscript:
    """validate_ghostscript のテスト"""

    def test_empty_path(self):
        """空パスは無効"""
        assert GhostscriptDetector.validate_ghostscript("") is False

    def test_none_path(self):
        """Noneは無効"""
        assert GhostscriptDetector.validate_ghostscript(None) is False

    def test_nonexistent_path(self):
        """存在しないパスは無効"""
        assert GhostscriptDetector.validate_ghostscript(r"C:\nonexistent\gs.exe") is False

    def test_valid_executable_name(self, tmp_path):
        """有効な実行ファイル名"""
        gs_exe = tmp_path / "gswin64c.exe"
        gs_exe.write_text("fake")
        assert GhostscriptDetector.validate_ghostscript(str(gs_exe)) is True

    def test_invalid_executable_name(self, tmp_path):
        """無効な実行ファイル名"""
        wrong_exe = tmp_path / "notepad.exe"
        wrong_exe.write_text("fake")
        assert GhostscriptDetector.validate_ghostscript(str(wrong_exe)) is False

    def test_directory_not_file(self, tmp_path):
        """ディレクトリはファイルではないので無効"""
        assert GhostscriptDetector.validate_ghostscript(str(tmp_path)) is False

    def test_all_valid_names(self, tmp_path):
        """全ての有効な実行ファイル名を検証"""
        for name in GhostscriptDetector.GS_EXECUTABLES:
            exe = tmp_path / name
            exe.write_text("fake")
            assert GhostscriptDetector.validate_ghostscript(str(exe)) is True
            exe.unlink()


@pytest.mark.unit
class TestVerify:
    """verify のテスト"""

    def test_empty_path(self):
        """空パスはFalse"""
        assert GhostscriptDetector.verify("") is False

    def test_nonexistent_path(self):
        """存在しないパスはFalse"""
        assert GhostscriptDetector.verify(r"C:\nonexistent\gs.exe") is False

    @patch("infrastructure.ghostscript.subprocess.run")
    def test_successful_verify(self, mock_run, tmp_path):
        """正常な動作確認"""
        gs_exe = tmp_path / "gswin64c.exe"
        gs_exe.write_text("fake")
        mock_run.return_value = MagicMock(returncode=0, stdout="10.02.0\n")

        assert GhostscriptDetector.verify(str(gs_exe)) is True
        mock_run.assert_called_once()

    @patch("infrastructure.ghostscript.subprocess.run")
    def test_failed_verify(self, mock_run, tmp_path):
        """動作確認失敗"""
        gs_exe = tmp_path / "gswin64c.exe"
        gs_exe.write_text("fake")
        mock_run.return_value = MagicMock(returncode=1)

        assert GhostscriptDetector.verify(str(gs_exe)) is False

    @patch("infrastructure.ghostscript.subprocess.run")
    def test_verify_exception(self, mock_run, tmp_path):
        """subprocess例外"""
        gs_exe = tmp_path / "gswin64c.exe"
        gs_exe.write_text("fake")
        mock_run.side_effect = OSError("connection refused")

        assert GhostscriptDetector.verify(str(gs_exe)) is False


@pytest.mark.unit
class TestParseVersion:
    """_parse_version のテスト"""

    def test_standard_version(self):
        assert GhostscriptDetector._parse_version("10.02.0") == (10, 2, 0)

    def test_two_part_version(self):
        assert GhostscriptDetector._parse_version("9.56") == (9, 56)

    def test_single_part(self):
        assert GhostscriptDetector._parse_version("10") == (10,)

    def test_invalid_version(self):
        assert GhostscriptDetector._parse_version("invalid") == (0,)

    def test_empty_string(self):
        assert GhostscriptDetector._parse_version("") == (0,)


@pytest.mark.unit
class TestCheckEnvironmentVariables:
    """_check_environment_variables のテスト"""

    @patch.dict(os.environ, {}, clear=True)
    def test_no_env_vars(self):
        """環境変数なし"""
        assert GhostscriptDetector._check_environment_variables() is None

    @patch.dict(os.environ, {"GS_DLL": r"C:\nonexistent\gsdll64.dll"})
    def test_gs_dll_nonexistent(self):
        """GS_DLLが存在しないパス"""
        assert GhostscriptDetector._check_environment_variables() is None

    def test_gs_dll_with_valid_exe(self, tmp_path):
        """GS_DLLから実行ファイルを検出"""
        dll = tmp_path / "gsdll64.dll"
        dll.write_text("fake")
        exe = tmp_path / "gswin64c.exe"
        exe.write_text("fake")

        with patch.dict(os.environ, {"GS_DLL": str(dll)}):
            result = GhostscriptDetector._check_environment_variables()
            assert result == str(exe)


@pytest.mark.unit
class TestCheckPathEnv:
    """_check_path_env のテスト"""

    def test_gs_in_path(self, tmp_path):
        """PATH上にGSが存在"""
        gs_exe = tmp_path / "gswin64c.exe"
        gs_exe.write_text("fake")

        with patch.dict(os.environ, {"PATH": str(tmp_path)}):
            result = GhostscriptDetector._check_path_env()
            assert result == str(gs_exe)

    def test_gs_not_in_path(self, tmp_path):
        """PATH上にGSが存在しない"""
        with patch.dict(os.environ, {"PATH": str(tmp_path)}):
            result = GhostscriptDetector._check_path_env()
            assert result is None


@pytest.mark.unit
class TestGetInstallInstructions:
    """get_install_instructions のテスト"""

    def test_returns_instructions(self):
        """インストール手順を返す"""
        instructions = GhostscriptDetector.get_install_instructions()
        assert "ghostscript.com" in instructions
        assert "ダウンロード" in instructions


@pytest.mark.unit
class TestDetect:
    """detect のテスト"""

    @patch.object(GhostscriptDetector, "_check_environment_variables", return_value=None)
    @patch.object(GhostscriptDetector, "_check_registry", return_value=None)
    @patch.object(GhostscriptDetector, "_check_standard_paths", return_value=None)
    @patch.object(GhostscriptDetector, "_check_path_env", return_value=None)
    def test_nothing_found(self, *mocks):
        """何も見つからない"""
        assert GhostscriptDetector.detect() is None

    @patch.object(GhostscriptDetector, "_check_environment_variables", return_value=r"C:\gs\bin\gs.exe")
    def test_env_var_priority(self, mock_env):
        """環境変数が最優先"""
        result = GhostscriptDetector.detect()
        assert result == r"C:\gs\bin\gs.exe"

    @patch.object(GhostscriptDetector, "_check_environment_variables", return_value=None)
    @patch.object(GhostscriptDetector, "_check_registry", return_value=r"C:\gs\bin\gs.exe")
    def test_registry_fallback(self, *mocks):
        """環境変数なし→レジストリ"""
        result = GhostscriptDetector.detect()
        assert result == r"C:\gs\bin\gs.exe"
