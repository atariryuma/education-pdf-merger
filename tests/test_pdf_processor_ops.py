"""
PDFProcessor のユニットテスト

PDF操作（マージ、分割、TOC、ページ番号、ブックマーク）をテスト
"""
import os
import pytest
from unittest.mock import MagicMock, patch

fitz = pytest.importorskip("fitz", reason="PyMuPDF not installed")
pytest.importorskip("reportlab", reason="reportlab not installed")

from shared.exceptions import PDFProcessingError
from core.pdf_processor import PDFProcessor


@pytest.fixture(autouse=True)
def _mock_pdfmetrics():
    """全テストでpdfmetricsをモック（フォント登録の副作用を防止）"""
    with patch('core.pdf_processor.pdfmetrics'):
        yield


@pytest.fixture
def mock_config():
    """ConfigLoaderモック"""
    config = MagicMock()
    config.get.return_value = "C:\\Windows\\Fonts\\msmincho.ttc"
    return config


@pytest.fixture
def real_pdf(temp_dir):
    """fitz で読める実際のPDFを作成"""
    pdf_path = os.path.join(temp_dir, "real.pdf")
    doc = fitz.open()
    for _ in range(3):
        doc.new_page()
    doc.save(pdf_path)
    doc.close()
    return pdf_path


@pytest.mark.unit
class TestMergePdfs:
    """merge_pdfs のテスト"""

    def test_merge_existing_pdfs(self, temp_dir, mock_config):
        """存在するPDFがマージされる"""
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

    def test_merge_skips_none_paths(self, temp_dir, mock_config):
        """Noneパスがスキップされる"""
        path = os.path.join(temp_dir, "test.pdf")
        doc = fitz.open()
        doc.new_page()
        doc.save(path)
        doc.close()

        processor = PDFProcessor(mock_config)
        output = os.path.join(temp_dir, "merged.pdf")
        processor.merge_pdfs([path, None, "/nonexistent.pdf"], output)

        assert os.path.exists(output)

    def test_merge_empty_list(self, temp_dir, mock_config):
        """空リストでもエラーにならない"""
        processor = PDFProcessor(mock_config)
        output = os.path.join(temp_dir, "merged.pdf")
        processor.merge_pdfs([], output)

        assert os.path.exists(output)
        # 結果が有効なPDFであることを確認
        with fitz.open(output) as doc:
            assert doc.page_count == 0


@pytest.mark.unit
class TestGetPageCount:
    """get_page_count のテスト"""

    def test_correct_page_count(self, real_pdf, mock_config):
        """正しいページ数が返される"""
        processor = PDFProcessor(mock_config)
        count = processor.get_page_count(real_pdf)
        assert count == 3

    def test_nonexistent_file_raises(self, mock_config):
        """存在しないファイルでPDFProcessingError"""
        processor = PDFProcessor(mock_config)
        with pytest.raises(PDFProcessingError, match="読み込みに失敗"):
            processor.get_page_count("/nonexistent.pdf")


@pytest.mark.unit
class TestSplitPdf:
    """split_pdf のテスト"""

    def test_split_creates_two_files(self, real_pdf, temp_dir, mock_config):
        """分割で表紙と残りの2ファイルが作成される"""
        processor = PDFProcessor(mock_config)
        cover, remainder = processor.split_pdf(real_pdf, temp_dir)

        assert os.path.exists(cover)
        assert os.path.exists(remainder)

        with fitz.open(cover) as doc:
            assert doc.page_count == 1
        with fitz.open(remainder) as doc:
            assert doc.page_count == 2  # 3ページ中残り2ページ

    def test_split_single_page_returns_none_remainder(self, temp_dir, mock_config):
        """1ページPDFの分割でremainderがNone"""
        single_pdf = os.path.join(temp_dir, "single.pdf")
        doc = fitz.open()
        doc.new_page()
        doc.save(single_pdf)
        doc.close()

        processor = PDFProcessor(mock_config)
        cover, remainder = processor.split_pdf(single_pdf, temp_dir)

        assert os.path.exists(cover)
        with fitz.open(cover) as doc:
            assert doc.page_count == 1
        assert remainder is None

    def test_split_nonexistent_raises(self, temp_dir, mock_config):
        """存在しないPDFでPDFProcessingError"""
        processor = PDFProcessor(mock_config)
        with pytest.raises(PDFProcessingError):
            processor.split_pdf("/nonexistent.pdf", temp_dir)


@pytest.mark.unit
class TestAddPageNumbers:
    """add_page_numbers のテスト"""

    def test_page_numbers_added(self, real_pdf, mock_config):
        """ページ番号が追加されてもファイルが壊れない"""
        processor = PDFProcessor(mock_config)
        processor.add_page_numbers(real_pdf, exclude_first_pages=1)

        with fitz.open(real_pdf) as doc:
            assert doc.page_count == 3  # ページ数は変わらない
            # exclude_first_pages=1なので、2ページ目以降にページ番号テキストが追加されている
            page2_text = doc[1].get_text()
            assert page2_text.strip() != ""  # ページ番号テキストが存在する


@pytest.mark.unit
class TestSetPdfOutlines:
    """set_pdf_outlines のテスト"""

    def test_outlines_set(self, real_pdf, mock_config):
        """アウトラインが設定される"""
        processor = PDFProcessor(mock_config)
        toc_entries = [
            ("Section 1", 1, 1),
            ("Subsection", 2, 2),
        ]
        processor.set_pdf_outlines(real_pdf, toc_entries)

        with fitz.open(real_pdf) as doc:
            toc = doc.get_toc()
            assert len(toc) == 2
            # アウトライン名とページ番号を検証
            assert toc[0][1] == "Section 1"
            assert toc[0][2] == 1
            assert toc[1][1] == "Subsection"
            assert toc[1][2] == 2

    def test_outlines_page_clamped(self, real_pdf, mock_config):
        """範囲外のページ番号が補正される"""
        processor = PDFProcessor(mock_config)
        toc_entries = [
            ("Over range", 1, 999),  # 3ページのPDFに999
        ]
        processor.set_pdf_outlines(real_pdf, toc_entries)

        with fitz.open(real_pdf) as doc:
            toc = doc.get_toc()
            assert len(toc) == 1
            assert toc[0][2] <= doc.page_count


@pytest.mark.unit
class TestCompressPdf:
    """compress_pdf のテスト"""

    @patch('core.pdf_processor.subprocess')
    def test_compress_success(self, mock_subprocess, real_pdf, mock_config):
        """圧縮成功でTrueが返る"""
        mock_subprocess.run.return_value = MagicMock(returncode=0)
        mock_subprocess.TimeoutExpired = TimeoutError
        mock_subprocess.CalledProcessError = Exception

        processor = PDFProcessor(mock_config)
        # _atomic_pdf_operation内でos.replaceが呼ばれるためモック
        with patch('core.pdf_processor.os.replace'):
            result = processor.compress_pdf(real_pdf)

        assert result is True

    @patch('core.pdf_processor.subprocess')
    def test_compress_timeout_returns_false(self, mock_subprocess, real_pdf, mock_config):
        """タイムアウトでFalseが返る"""
        import subprocess as real_subprocess

        mock_subprocess.TimeoutExpired = real_subprocess.TimeoutExpired
        mock_subprocess.CalledProcessError = real_subprocess.CalledProcessError
        mock_subprocess.run.side_effect = real_subprocess.TimeoutExpired("gs", 60)

        processor = PDFProcessor(mock_config)
        result = processor.compress_pdf(real_pdf)

        assert result is False
