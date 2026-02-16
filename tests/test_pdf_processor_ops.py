"""
PDFProcessor のユニットテスト

PDF操作（マージ、分割、TOC、ページ番号、ブックマーク）をテスト
"""
import os
import pytest
from unittest.mock import MagicMock, patch, mock_open

from exceptions import PDFProcessingError


@pytest.fixture
def mock_config():
    """ConfigLoaderモック"""
    config = MagicMock()
    config.get.return_value = "C:\\Windows\\Fonts\\msmincho.ttc"
    return config


@pytest.fixture
def real_pdf(temp_dir):
    """fitz で読める実際のPDFを作成"""
    try:
        import fitz
    except ImportError:
        pytest.skip("PyMuPDF not installed")

    pdf_path = os.path.join(temp_dir, "real.pdf")
    doc = fitz.open()
    for _ in range(3):
        doc.new_page()
    doc.save(pdf_path)
    doc.close()
    return pdf_path


class TestMergePdfs:
    """merge_pdfs のテスト"""

    @patch('pdf_processor.pdfmetrics')
    def test_merge_existing_pdfs(self, mock_metrics, temp_dir, mock_config):
        """存在するPDFがマージされる"""
        from pdf_processor import PDFProcessor

        try:
            import fitz
        except ImportError:
            pytest.skip("PyMuPDF not installed")

        # テスト用PDF作成
        pdfs = []
        for i in range(3):
            path = os.path.join(temp_dir, f"test{i}.pdf")
            doc = fitz.open()
            doc.new_page()
            doc.save(path)
            doc.close()
            pdfs.append(path)

        processor = PDFProcessor(mock_config)
        output = os.path.join(temp_dir, "merged.pdf")
        processor.merge_pdfs(pdfs, output)

        assert os.path.exists(output)
        with fitz.open(output) as merged:
            assert merged.page_count == 3

    @patch('pdf_processor.pdfmetrics')
    def test_merge_skips_none_paths(self, mock_metrics, temp_dir, mock_config):
        """Noneパスがスキップされる"""
        from pdf_processor import PDFProcessor

        try:
            import fitz
        except ImportError:
            pytest.skip("PyMuPDF not installed")

        path = os.path.join(temp_dir, "test.pdf")
        doc = fitz.open()
        doc.new_page()
        doc.save(path)
        doc.close()

        processor = PDFProcessor(mock_config)
        output = os.path.join(temp_dir, "merged.pdf")
        processor.merge_pdfs([path, None, "/nonexistent.pdf"], output)

        assert os.path.exists(output)

    @patch('pdf_processor.pdfmetrics')
    def test_merge_empty_list(self, mock_metrics, temp_dir, mock_config):
        """空リストでもエラーにならない"""
        from pdf_processor import PDFProcessor

        processor = PDFProcessor(mock_config)
        output = os.path.join(temp_dir, "merged.pdf")
        processor.merge_pdfs([], output)

        assert os.path.exists(output)


class TestGetPageCount:
    """get_page_count のテスト"""

    @patch('pdf_processor.pdfmetrics')
    def test_correct_page_count(self, mock_metrics, real_pdf, mock_config):
        """正しいページ数が返される"""
        from pdf_processor import PDFProcessor

        processor = PDFProcessor(mock_config)
        count = processor.get_page_count(real_pdf)
        assert count == 3

    @patch('pdf_processor.pdfmetrics')
    def test_nonexistent_file_raises(self, mock_metrics, mock_config):
        """存在しないファイルでPDFProcessingError"""
        from pdf_processor import PDFProcessor

        processor = PDFProcessor(mock_config)
        with pytest.raises(PDFProcessingError, match="読み込みに失敗"):
            processor.get_page_count("/nonexistent.pdf")


class TestSplitPdf:
    """split_pdf のテスト"""

    @patch('pdf_processor.pdfmetrics')
    def test_split_creates_two_files(self, mock_metrics, real_pdf, temp_dir, mock_config):
        """分割で表紙と残りの2ファイルが作成される"""
        from pdf_processor import PDFProcessor

        processor = PDFProcessor(mock_config)
        cover, remainder = processor.split_pdf(real_pdf, temp_dir)

        assert os.path.exists(cover)
        assert os.path.exists(remainder)

        import fitz
        with fitz.open(cover) as doc:
            assert doc.page_count == 1
        with fitz.open(remainder) as doc:
            assert doc.page_count == 2  # 3ページ中残り2ページ

    @patch('pdf_processor.pdfmetrics')
    def test_split_nonexistent_raises(self, mock_metrics, temp_dir, mock_config):
        """存在しないPDFでPDFProcessingError"""
        from pdf_processor import PDFProcessor

        processor = PDFProcessor(mock_config)
        with pytest.raises(PDFProcessingError):
            processor.split_pdf("/nonexistent.pdf", temp_dir)


class TestAddPageNumbers:
    """add_page_numbers のテスト"""

    @patch('pdf_processor.pdfmetrics')
    def test_page_numbers_added(self, mock_metrics, real_pdf, mock_config):
        """ページ番号が追加されてもファイルが壊れない"""
        from pdf_processor import PDFProcessor
        import fitz

        processor = PDFProcessor(mock_config)
        processor.add_page_numbers(real_pdf, exclude_first_pages=1)

        with fitz.open(real_pdf) as doc:
            assert doc.page_count == 3  # ページ数は変わらない


class TestSetPdfOutlines:
    """set_pdf_outlines のテスト"""

    @patch('pdf_processor.pdfmetrics')
    def test_outlines_set(self, mock_metrics, real_pdf, mock_config):
        """アウトラインが設定される"""
        from pdf_processor import PDFProcessor
        import fitz

        processor = PDFProcessor(mock_config)
        toc_entries = [
            ("Section 1", 1, 1),
            ("Subsection", 2, 2),
        ]
        processor.set_pdf_outlines(real_pdf, toc_entries)

        with fitz.open(real_pdf) as doc:
            toc = doc.get_toc()
            assert len(toc) == 2

    @patch('pdf_processor.pdfmetrics')
    def test_outlines_page_clamped(self, mock_metrics, real_pdf, mock_config):
        """範囲外のページ番号が補正される"""
        from pdf_processor import PDFProcessor
        import fitz

        processor = PDFProcessor(mock_config)
        toc_entries = [
            ("Over range", 1, 999),  # 3ページのPDFに999
        ]
        processor.set_pdf_outlines(real_pdf, toc_entries)

        with fitz.open(real_pdf) as doc:
            toc = doc.get_toc()
            assert len(toc) == 1
            assert toc[0][2] <= doc.page_count


class TestCompressPdf:
    """compress_pdf のテスト"""

    @patch('pdf_processor.pdfmetrics')
    @patch('pdf_processor.subprocess')
    def test_compress_success(self, mock_subprocess, mock_metrics, real_pdf, mock_config):
        """圧縮成功でTrueが返る"""
        from pdf_processor import PDFProcessor

        mock_subprocess.run.return_value = MagicMock(returncode=0)
        mock_subprocess.TimeoutExpired = TimeoutError
        mock_subprocess.CalledProcessError = Exception

        processor = PDFProcessor(mock_config)
        # _atomic_pdf_operation内でos.replaceが呼ばれるためモック
        with patch('pdf_processor.os.replace'):
            result = processor.compress_pdf(real_pdf)

        assert result is True

    @patch('pdf_processor.pdfmetrics')
    @patch('pdf_processor.subprocess')
    def test_compress_timeout_returns_false(self, mock_subprocess, mock_metrics, real_pdf, mock_config):
        """タイムアウトでFalseが返る"""
        import subprocess as real_subprocess
        from pdf_processor import PDFProcessor

        mock_subprocess.TimeoutExpired = real_subprocess.TimeoutExpired
        mock_subprocess.CalledProcessError = real_subprocess.CalledProcessError
        mock_subprocess.run.side_effect = real_subprocess.TimeoutExpired("gs", 60)

        processor = PDFProcessor(mock_config)
        result = processor.compress_pdf(real_pdf)

        assert result is False
