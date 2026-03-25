"""
統合テスト - PDF統合の全フロー検証

実際のファイルを使用した統合テストを実施
"""
import os
import tempfile
import shutil
from pathlib import Path
from typing import Generator
import pytest


def _has_win32com() -> bool:
    """Check if win32com.client is available."""
    try:
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        return False


from infrastructure.config_loader import ConfigLoader
from core.pdf_converter import PDFConverter
from core.pdf_processor import PDFProcessor
from core.document_collector import DocumentCollector
from core.pdf_merge_orchestrator import PDFMergeOrchestrator


@pytest.fixture
def temp_workspace() -> Generator[Path, None, None]:
    """
    一時作業ディレクトリを作成

    Yields:
        Path: 一時ディレクトリのパス
    """
    temp_dir = Path(tempfile.mkdtemp(prefix="pdf_merge_test_"))
    try:
        yield temp_dir
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


@pytest.fixture
def mock_config(temp_workspace: Path) -> ConfigLoader:
    """
    テスト用のConfigLoaderを作成

    Args:
        temp_workspace: 一時ディレクトリ

    Returns:
        ConfigLoader: テスト用設定
    """
    config_file = temp_workspace / "test_config.json"
    config_data = {
        "year": "2025",
        "year_short": "R7",
        "base_paths": {
            "local_temp": str(temp_workspace / "temp")
        },
        "fonts": {
            "mincho": "C:\\Windows\\Fonts\\msgothic.ttc"  # システムフォント
        },
        "ghostscript": {
            "executable": "gswin64c.exe"
        }
    }

    import json
    with open(config_file, 'w', encoding='utf-8') as f:
        json.dump(config_data, f, ensure_ascii=False, indent=2)

    return ConfigLoader(str(config_file))


@pytest.fixture
def sample_input_dir(temp_workspace: Path) -> Path:
    """
    サンプル入力ディレクトリを作成

    Args:
        temp_workspace: 一時ディレクトリ

    Returns:
        Path: サンプル入力ディレクトリ
    """
    input_dir = temp_workspace / "input"
    input_dir.mkdir(parents=True)

    # ダミーPDFバイト列（1ページの最小PDF）
    dummy_pdf = (
        b'%PDF-1.4\n%\xe2\xe3\xcf\xd3\n'
        b'1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n'
        b'2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n'
        b'3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] >>\nendobj\n'
        b'xref\n0 4\n0000000000 65535 f \n'
        b'0000000015 00000 n \n0000000068 00000 n \n0000000131 00000 n \n'
        b'trailer\n<< /Size 4 /Root 1 0 R >>\nstartxref\n214\n%%EOF\n'
    )

    # 表紙ファイル（ダミーPDF）
    (input_dir / "表紙.pdf").write_bytes(dummy_pdf)

    # メインディレクトリ1
    main_dir1 = input_dir / "01_国語"
    main_dir1.mkdir()
    (main_dir1 / "01_1年.pdf").write_bytes(dummy_pdf)
    (main_dir1 / "02_2年.pdf").write_bytes(dummy_pdf)

    # メインディレクトリ2
    main_dir2 = input_dir / "02_算数"
    main_dir2.mkdir()
    (main_dir2 / "01_1年.pdf").write_bytes(dummy_pdf)

    return input_dir


@pytest.mark.integration
class TestPDFMergeIntegration:
    """PDF統合の統合テスト"""

    def test_config_loader_initialization(self, mock_config: ConfigLoader):
        """ConfigLoaderが正しく初期化されることを確認"""
        assert mock_config.year == "2025"
        assert mock_config.year_short == "R7"
        assert mock_config.get('year') == "2025"

    def test_temp_dir_creation(self, mock_config: ConfigLoader):
        """一時ディレクトリが正しく作成されることを確認"""
        temp_dir = mock_config.get_temp_dir()
        assert os.path.exists(temp_dir)
        assert os.path.isdir(temp_dir)

    def test_pdf_processor_initialization(self, mock_config: ConfigLoader):
        """PDFProcessorが正しく初期化されることを確認"""
        processor = PDFProcessor(mock_config)
        assert processor.config == mock_config

    def test_pdf_converter_initialization(self, mock_config: ConfigLoader):
        """PDFConverterが正しく初期化されることを確認"""
        temp_dir = mock_config.get_temp_dir()
        converter = PDFConverter(temp_dir, config=mock_config)
        assert converter.temp_dir == temp_dir
        assert converter.config == mock_config

    def test_document_collector_initialization(self, mock_config: ConfigLoader):
        """DocumentCollectorが正しく初期化されることを確認"""
        temp_dir = mock_config.get_temp_dir()
        converter = PDFConverter(temp_dir, config=mock_config)
        processor = PDFProcessor(mock_config)
        collector = DocumentCollector(converter, processor)
        assert collector.converter == converter
        assert collector.processor == processor

    def test_orchestrator_initialization(self, mock_config: ConfigLoader):
        """PDFMergeOrchestratorが正しく初期化されることを確認"""
        temp_dir = mock_config.get_temp_dir()
        converter = PDFConverter(temp_dir, config=mock_config)
        processor = PDFProcessor(mock_config)
        collector = DocumentCollector(converter, processor)
        orchestrator = PDFMergeOrchestrator(mock_config, converter, processor, collector)

        assert orchestrator.config == mock_config
        assert orchestrator.converter == converter
        assert orchestrator.processor == processor
        assert orchestrator.collector == collector

    def test_sample_directory_structure(self, sample_input_dir: Path):
        """サンプルディレクトリ構造が正しく作成されることを確認"""
        assert sample_input_dir.exists()
        assert (sample_input_dir / "表紙.pdf").exists()
        assert (sample_input_dir / "01_国語").is_dir()
        assert (sample_input_dir / "02_算数").is_dir()

        # サブディレクトリ内のファイルも正しく作成されている
        kokugo_files = list((sample_input_dir / "01_国語").iterdir())
        assert len(kokugo_files) == 2
        sansu_files = list((sample_input_dir / "02_算数").iterdir())
        assert len(sansu_files) == 1

    @pytest.mark.skipif(
        not _has_win32com(),
        reason="win32com.client not available (Office not installed)"
    )
    def test_full_pdf_merge_flow(
        self,
        mock_config: ConfigLoader,
        sample_input_dir: Path,
        temp_workspace: Path
    ):
        """
        完全なPDF統合フローのテスト（外部依存のため通常はスキップ）

        Args:
            mock_config: テスト用設定
            sample_input_dir: サンプル入力ディレクトリ
            temp_workspace: 一時作業ディレクトリ
        """
        output_pdf = temp_workspace / "output.pdf"

        # コンポーネント初期化
        temp_dir = mock_config.get_temp_dir()
        converter = PDFConverter(temp_dir, config=mock_config)
        processor = PDFProcessor(mock_config)
        collector = DocumentCollector(converter, processor)
        orchestrator = PDFMergeOrchestrator(mock_config, converter, processor, collector)

        # PDF統合実行
        orchestrator.create_merged_pdf(
            target_dir=str(sample_input_dir),
            output_pdf=str(output_pdf),
            create_separator_for_subfolder=True
        )

        # 結果検証
        assert output_pdf.exists()
        assert output_pdf.stat().st_size > 0


@pytest.mark.integration
class TestPathValidator:
    """PathValidatorの統合テスト"""

    def test_validate_existing_directory(self, temp_workspace: Path):
        """存在するディレクトリの検証"""
        from infrastructure.path_validator import PathValidator
        import os

        is_valid, error_msg, validated_path = PathValidator.validate_directory(
            str(temp_workspace),
            must_exist=True
        )

        assert is_valid is True
        assert error_msg is None
        # Windows short path (RUNNER~1) vs long path (runneradmin) の違いを吸収
        # os.path.realpath() は短縮パスを長形式に変換する
        assert os.path.realpath(str(validated_path)) == os.path.realpath(str(temp_workspace))

    def test_validate_nonexistent_directory(self, temp_workspace: Path):
        """存在しないディレクトリの検証"""
        from infrastructure.path_validator import PathValidator

        nonexistent = temp_workspace / "nonexistent"
        is_valid, error_msg, validated_path = PathValidator.validate_directory(
            str(nonexistent),
            must_exist=True
        )

        assert is_valid is False
        assert error_msg is not None
        assert "存在しません" in error_msg

    def test_sanitize_filename(self):
        """ファイル名のサニタイズ"""
        from infrastructure.path_validator import PathValidator

        # 特殊文字を含むファイル名
        unsafe_name = "test<>file:name?.txt"
        safe_name = PathValidator.sanitize_filename(unsafe_name)

        assert "<" not in safe_name
        assert ">" not in safe_name
        assert ":" not in safe_name
        assert "?" not in safe_name

    def test_sanitize_windows_reserved_name(self):
        """Windows予約名のサニタイズ"""
        from infrastructure.path_validator import PathValidator

        reserved_name = "CON.txt"
        safe_name = PathValidator.sanitize_filename(reserved_name)

        # 予約名の回避を確認
        assert not safe_name.upper().startswith("CON")


@pytest.mark.integration
class TestExceptionHandling:
    """例外処理の統合テスト"""

    def test_configuration_error_chain(self, temp_workspace: Path):
        """ConfigurationErrorの例外チェーン"""
        from infrastructure.config_loader import ConfigLoader
        from shared.exceptions import ConfigurationError

        nonexistent_config = temp_workspace / "nonexistent_config.json"

        with pytest.raises(ConfigurationError) as exc_info:
            ConfigLoader(str(nonexistent_config))

        # 例外チェーンの確認
        assert exc_info.value.original_error is not None
        assert isinstance(exc_info.value.original_error, FileNotFoundError)

    def test_cancelled_error(self):
        """CancelledErrorの動作確認"""
        from shared.exceptions import CancelledError

        cancel_flag = False

        def cancel_check():
            return cancel_flag

        # キャンセルされていない状態
        assert cancel_check() is False

        # キャンセル状態に変更
        cancel_flag = True
        assert cancel_check() is True

        # CancelledErrorの生成
        error = CancelledError("テストキャンセル")
        assert "テストキャンセル" in str(error)


if __name__ == "__main__":
    pytest.main([__file__, "-v", "--tb=short"])
