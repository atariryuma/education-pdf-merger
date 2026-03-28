"""
PathValidator のユニットテスト

パス検証、サニタイズ、セキュリティチェックをテスト
"""

import pytest
from pathlib import Path

from infrastructure.path_validator import (
    PathValidator,
    PathValidationError,
    _check_path_security,
)


@pytest.mark.unit
class TestNormalizePath:
    """normalize_path のテスト"""

    def test_valid_absolute_path(self, tmp_path):
        """有効な絶対パスを正規化"""
        result = PathValidator.normalize_path(str(tmp_path))
        assert result == tmp_path.resolve()

    def test_strips_whitespace(self, tmp_path):
        """前後の空白を除去"""
        result = PathValidator.normalize_path(f"  {tmp_path}  ")
        assert result == tmp_path.resolve()

    def test_removes_newlines_and_tabs(self, tmp_path):
        """改行・タブを除去"""
        result = PathValidator.normalize_path(f"{tmp_path}\n\r\t")
        assert result == tmp_path.resolve()

    def test_empty_string_raises(self):
        """空文字列でエラー"""
        with pytest.raises(PathValidationError, match="パスが空です"):
            PathValidator.normalize_path("")

    def test_whitespace_only_raises(self):
        """空白のみでエラー"""
        with pytest.raises(PathValidationError, match="パスが空です"):
            PathValidator.normalize_path("   ")

    def test_none_raises(self):
        """Noneでエラー"""
        with pytest.raises(PathValidationError):
            PathValidator.normalize_path(None)


@pytest.mark.unit
class TestValidateDirectory:
    """validate_directory のテスト"""

    def test_valid_existing_directory(self, tmp_path):
        """存在するディレクトリ"""
        is_valid, error_msg, path = PathValidator.validate_directory(str(tmp_path))
        assert is_valid is True
        assert error_msg is None
        assert path == tmp_path.resolve()

    def test_nonexistent_directory(self, tmp_path):
        """存在しないディレクトリ"""
        is_valid, error_msg, path = PathValidator.validate_directory(
            str(tmp_path / "nonexistent"), must_exist=True
        )
        assert is_valid is False
        assert "存在しません" in error_msg
        assert "親ディレクトリは存在します" in error_msg
        assert path is None

    def test_nonexistent_with_missing_parent(self):
        """親ディレクトリも存在しない"""
        is_valid, error_msg, path = PathValidator.validate_directory(
            r"Z:\nonexistent\deeply\nested", must_exist=True
        )
        assert is_valid is False
        assert path is None

    def test_file_not_directory(self, tmp_path):
        """ファイルをディレクトリとして検証"""
        file = tmp_path / "test.txt"
        file.write_text("test")
        is_valid, error_msg, path = PathValidator.validate_directory(str(file))
        assert is_valid is False
        assert "ディレクトリではありません" in error_msg

    def test_must_exist_false(self, tmp_path):
        """存在チェックなし"""
        is_valid, error_msg, path = PathValidator.validate_directory(
            str(tmp_path / "new_dir"), must_exist=False
        )
        assert is_valid is True
        assert path is not None

    def test_empty_path(self):
        """空パス"""
        is_valid, error_msg, path = PathValidator.validate_directory("")
        assert is_valid is False
        assert "パスが空です" in error_msg

    def test_base_dir_security_pass(self, tmp_path):
        """基準ディレクトリ配下のパス（通過）"""
        sub_dir = tmp_path / "sub"
        sub_dir.mkdir()
        is_valid, _, path = PathValidator.validate_directory(
            str(sub_dir), base_dir=tmp_path
        )
        assert is_valid is True

    def test_base_dir_security_fail(self, tmp_path):
        """基準ディレクトリ外のパス（拒否）"""
        other_dir = tmp_path.parent
        is_valid, error_msg, path = PathValidator.validate_directory(
            str(other_dir), base_dir=tmp_path
        )
        assert is_valid is False
        assert "許可されていない" in error_msg


@pytest.mark.unit
class TestValidateFilePath:
    """validate_file_path のテスト"""

    def test_valid_existing_file(self, tmp_path):
        """存在するファイル"""
        file = tmp_path / "test.pdf"
        file.write_text("test")
        is_valid, error_msg, path = PathValidator.validate_file_path(
            str(file), must_exist=True
        )
        assert is_valid is True
        assert error_msg is None

    def test_nonexistent_file(self, tmp_path):
        """存在しないファイル"""
        is_valid, error_msg, path = PathValidator.validate_file_path(
            str(tmp_path / "missing.pdf"), must_exist=True
        )
        assert is_valid is False
        assert "存在しません" in error_msg

    def test_allowed_extensions_pass(self, tmp_path):
        """許可された拡張子"""
        file = tmp_path / "test.pdf"
        file.write_text("test")
        is_valid, _, _ = PathValidator.validate_file_path(
            str(file), must_exist=True, allowed_extensions=[".pdf"]
        )
        assert is_valid is True

    def test_allowed_extensions_fail(self, tmp_path):
        """許可されていない拡張子"""
        file = tmp_path / "test.txt"
        file.write_text("test")
        is_valid, error_msg, _ = PathValidator.validate_file_path(
            str(file), must_exist=True, allowed_extensions=[".pdf"]
        )
        assert is_valid is False
        assert "許可されていない拡張子" in error_msg

    def test_extension_case_insensitive(self, tmp_path):
        """拡張子の大文字小文字を区別しない"""
        file = tmp_path / "test.PDF"
        file.write_text("test")
        is_valid, _, _ = PathValidator.validate_file_path(
            str(file), must_exist=True, allowed_extensions=[".pdf"]
        )
        assert is_valid is True

    def test_save_path_checks_parent(self, tmp_path):
        """保存先の親ディレクトリ存在チェック"""
        is_valid, _, path = PathValidator.validate_file_path(
            str(tmp_path / "output.pdf"), must_exist=False
        )
        assert is_valid is True

    def test_save_path_missing_parent(self):
        """保存先の親ディレクトリが存在しない"""
        is_valid, error_msg, _ = PathValidator.validate_file_path(
            r"Z:\nonexistent\output.pdf", must_exist=False
        )
        assert is_valid is False
        assert "保存先ディレクトリが存在しません" in error_msg

    def test_directory_as_file(self, tmp_path):
        """ディレクトリをファイルとして検証"""
        is_valid, error_msg, _ = PathValidator.validate_file_path(
            str(tmp_path), must_exist=True
        )
        assert is_valid is False
        assert "ファイルではありません" in error_msg


@pytest.mark.unit
class TestGetSafeInitialDir:
    """get_safe_initial_dir のテスト"""

    def test_returns_existing_directory(self, tmp_path):
        """既存ディレクトリをそのまま返す"""
        result = PathValidator.get_safe_initial_dir(str(tmp_path))
        assert result == tmp_path.resolve()

    def test_returns_parent_for_file(self, tmp_path):
        """ファイルの場合は親ディレクトリを返す"""
        file = tmp_path / "test.txt"
        file.write_text("test")
        result = PathValidator.get_safe_initial_dir(str(file))
        assert result == tmp_path.resolve()

    def test_returns_fallback_for_invalid(self, tmp_path):
        """無効なパスの場合はフォールバック"""
        result = PathValidator.get_safe_initial_dir(
            r"Z:\nonexistent", fallback=tmp_path
        )
        assert result == tmp_path

    def test_returns_home_for_empty(self):
        """空文字列の場合はホームディレクトリ"""
        result = PathValidator.get_safe_initial_dir("")
        assert result == Path.home()


@pytest.mark.unit
class TestSanitizeFilename:
    """sanitize_filename のテスト"""

    def test_valid_filename_unchanged(self):
        """有効なファイル名はそのまま"""
        assert PathValidator.sanitize_filename("document.pdf") == "document.pdf"

    def test_removes_invalid_chars(self):
        """無効な文字を置換"""
        result = PathValidator.sanitize_filename('file<>:"/\\|?*.pdf')
        assert "<" not in result
        assert ">" not in result
        assert ":" not in result
        assert '"' not in result
        assert "?" not in result
        assert "*" not in result

    def test_removes_control_chars(self):
        """制御文字を除去"""
        result = PathValidator.sanitize_filename("file\x00\x01name.txt")
        assert "\x00" not in result
        assert "\x01" not in result

    def test_windows_reserved_names(self):
        """Windows予約名を回避"""
        result = PathValidator.sanitize_filename("CON.txt")
        assert result != "CON.txt"
        assert result.endswith(".txt")

    def test_windows_reserved_case_insensitive(self):
        """予約名の大文字小文字を区別しない"""
        result = PathValidator.sanitize_filename("con.txt")
        assert not result.upper().startswith("CON.")

    def test_strips_leading_trailing(self):
        """先頭末尾のスペース・ドット・アンダースコアを削除"""
        assert PathValidator.sanitize_filename("...file.txt") == "file.txt"
        assert PathValidator.sanitize_filename("___file.txt") == "file.txt"

    def test_empty_returns_default(self):
        """空文字列でデフォルト名"""
        from shared.constants import PathConstants

        assert PathValidator.sanitize_filename("") == PathConstants.DEFAULT_FILENAME

    def test_only_invalid_chars_returns_default(self):
        """無効文字のみでデフォルト名"""
        from shared.constants import PathConstants

        result = PathValidator.sanitize_filename("***")
        assert result == PathConstants.DEFAULT_FILENAME

    def test_max_length(self):
        """最大文字数制限"""
        long_name = "a" * 300 + ".pdf"
        result = PathValidator.sanitize_filename(long_name, max_length=255)
        assert len(result) <= 255
        assert result.endswith(".pdf")

    def test_unicode_normalization(self):
        """Unicode正規化（NFD→NFC）"""
        # NFD形式の「が」(か + 濁点)
        nfd = "ファイル\u304b\u3099.txt"
        result = PathValidator.sanitize_filename(nfd)
        assert "が" in result or "\u304b\u3099" not in result

    def test_consecutive_replacements_collapsed(self):
        """連続する置換文字を統合"""
        result = PathValidator.sanitize_filename("a<>:<>.txt")
        assert "__" not in result  # 連続 _ が統合される


@pytest.mark.unit
class TestCheckPathSecurity:
    """_check_path_security のテスト"""

    def test_child_path_allowed(self, tmp_path):
        """子パスは許可"""
        child = (tmp_path / "sub").resolve()
        assert _check_path_security(child, tmp_path.resolve()) is True

    def test_parent_path_denied(self, tmp_path):
        """親パスは拒否"""
        child = tmp_path / "sub"
        child.mkdir()
        assert _check_path_security(tmp_path.resolve(), child.resolve()) is False

    def test_same_path_allowed(self, tmp_path):
        """同じパスは許可"""
        assert _check_path_security(tmp_path.resolve(), tmp_path.resolve()) is True
