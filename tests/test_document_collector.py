"""
DocumentCollector のユニットテスト

ディレクトリ走査・ファイル収集・目次構造生成をテスト
"""
import os
import pytest
from unittest.mock import MagicMock, patch

from document_collector import DocumentCollector
from exceptions import CancelledError, PDFProcessingError
from constants import PDFConstants


@pytest.fixture
def mock_converter():
    """PDFConverterモック"""
    converter = MagicMock()
    converter.convert.return_value = "/tmp/converted.pdf"
    converter.create_separator_page.return_value = "/tmp/separator.pdf"
    return converter


@pytest.fixture
def mock_processor():
    """PDFProcessorモック"""
    processor = MagicMock()
    processor.get_page_count.return_value = 2
    return processor


@pytest.fixture
def collector(mock_converter, mock_processor):
    """テスト用DocumentCollectorインスタンス"""
    return DocumentCollector(mock_converter, mock_processor)


class TestSanitizeName:
    """_sanitize_name のテスト"""

    def test_removes_leading_numbers(self):
        assert DocumentCollector._sanitize_name("01 教育計画") == "教育計画"

    def test_removes_leading_numbers_and_spaces(self):
        assert DocumentCollector._sanitize_name("123  テスト") == "テスト"

    def test_removes_underscores(self):
        assert DocumentCollector._sanitize_name("01_教育_計画") == "教育計画"

    def test_no_change_when_no_prefix(self):
        assert DocumentCollector._sanitize_name("教育計画") == "教育計画"

    def test_empty_string(self):
        assert DocumentCollector._sanitize_name("") == ""

    def test_only_numbers(self):
        assert DocumentCollector._sanitize_name("123") == ""


class TestCollectDocuments:
    """collect_documents のテスト"""

    def test_empty_directory_raises(self, collector, temp_dir, mock_converter):
        """空ディレクトリでPDFProcessingError"""
        mock_converter.convert.return_value = None
        mock_converter.create_separator_page.return_value = None

        # 空のサブディレクトリを作成（ファイルなし）
        os.makedirs(os.path.join(temp_dir, "subdir"))

        with pytest.raises(PDFProcessingError, match="処理可能なドキュメント"):
            collector.collect_documents(temp_dir)

    def test_root_files_collected(self, collector, temp_dir, mock_converter, mock_processor):
        """ルート直下のファイルが収集される"""
        # テストファイル作成
        test_file = os.path.join(temp_dir, "test.docx")
        with open(test_file, 'w') as f:
            f.write("dummy")

        mock_converter.convert.return_value = "/tmp/converted.pdf"
        mock_processor.get_page_count.return_value = 3

        toc_entries, content_pdfs = collector.collect_documents(temp_dir)

        assert len(content_pdfs) >= 1
        mock_converter.convert.assert_called()

    def test_cover_file_processed_separately(self, collector, temp_dir, mock_converter, mock_processor):
        """表紙ファイルはTOCに含まれずPDFにのみ追加される"""
        # 表紙ファイルとその他のファイル
        cover_file = os.path.join(temp_dir, "表紙.docx")
        other_file = os.path.join(temp_dir, "other.docx")
        with open(cover_file, 'w') as f:
            f.write("cover")
        with open(other_file, 'w') as f:
            f.write("other")

        mock_converter.convert.return_value = "/tmp/converted.pdf"
        mock_processor.get_page_count.return_value = 1

        toc_entries, content_pdfs = collector.collect_documents(temp_dir)

        # 表紙はTOCエントリに含まれない
        toc_names = [entry[0] for entry in toc_entries]
        assert "表紙" not in toc_names
        # 表紙もconvertされるが、content_pdfsには含まれる
        assert len(content_pdfs) >= 2
        # 表紙がconvertの呼び出しに含まれる
        all_convert_args = [call[0][0] for call in mock_converter.convert.call_args_list]
        assert any("表紙" in arg for arg in all_convert_args)

    def test_cancel_during_collection(self, mock_converter, mock_processor, temp_dir):
        """キャンセルチェックが機能する"""
        collector = DocumentCollector(
            mock_converter, mock_processor,
            cancel_check=lambda: True
        )

        test_file = os.path.join(temp_dir, "test.docx")
        with open(test_file, 'w') as f:
            f.write("dummy")

        with pytest.raises(CancelledError):
            collector.collect_documents(temp_dir)

    def test_directory_with_subdirs(self, collector, temp_dir, mock_converter, mock_processor):
        """サブディレクトリが大見出しとして処理される"""
        # サブディレクトリ作成
        sub_dir = os.path.join(temp_dir, "01 教育計画")
        os.makedirs(sub_dir)
        sub_file = os.path.join(sub_dir, "test.docx")
        with open(sub_file, 'w') as f:
            f.write("dummy")

        mock_converter.convert.return_value = "/tmp/converted.pdf"
        mock_converter.create_separator_page.return_value = "/tmp/separator.pdf"
        mock_processor.get_page_count.return_value = 1

        toc_entries, content_pdfs = collector.collect_documents(temp_dir)

        # 区切りページが作成される
        mock_converter.create_separator_page.assert_called()
        assert len(content_pdfs) > 0


class TestConvertAndAddPdf:
    """_convert_and_add_pdf のテスト"""

    def test_successful_conversion(self, collector, temp_dir, mock_converter, mock_processor):
        """変換成功時にページ数が加算される"""
        test_file = os.path.join(temp_dir, "test.docx")
        with open(test_file, 'w') as f:
            f.write("dummy")

        mock_converter.convert.return_value = "/tmp/out.pdf"
        mock_processor.get_page_count.return_value = 5

        content_pdfs = []
        result = collector._convert_and_add_pdf(test_file, content_pdfs, 10)

        assert result == 15  # 10 + 5ページ
        assert len(content_pdfs) == 1

    def test_failed_conversion(self, collector, temp_dir, mock_converter):
        """変換失敗時はページ数が変わらない"""
        test_file = os.path.join(temp_dir, "test.bad")
        with open(test_file, 'w') as f:
            f.write("dummy")

        mock_converter.convert.return_value = None

        content_pdfs = []
        result = collector._convert_and_add_pdf(test_file, content_pdfs, 10)

        assert result == 10  # 変更なし
        assert len(content_pdfs) == 0


class TestProcessRootFile:
    """_process_root_file のテスト"""

    def test_adds_toc_entry(self, collector, temp_dir, mock_converter, mock_processor):
        """TOCエントリが追加される"""
        test_file = os.path.join(temp_dir, "01 概要.docx")
        with open(test_file, 'w') as f:
            f.write("dummy")

        mock_converter.convert.return_value = "/tmp/out.pdf"
        mock_processor.get_page_count.return_value = 3

        toc_entries = []
        content_pdfs = []
        result = collector._process_root_file(
            test_file, "01 概要.docx", toc_entries, content_pdfs, 1
        )

        assert result == 4  # 1 + 3ページ
        assert len(toc_entries) == 1
        assert toc_entries[0][0] == "概要"  # サニタイズ済み名前
        assert toc_entries[0][1] == PDFConstants.HEADING_LEVEL_SUB
