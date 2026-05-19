"""
DocumentCollector のユニットテスト

ディレクトリ走査・ファイル収集・目次構造生成をテスト
"""
import os
import pytest
from unittest.mock import MagicMock

from core.document_collector import DocumentCollector
from shared.exceptions import CancelledError, PDFProcessingError
from shared.constants import PDFConstants


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


@pytest.mark.unit
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


@pytest.mark.unit
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

    def test_root_files_collected(
        self, collector, temp_dir, mock_converter, mock_processor
    ):
        """ルート直下のファイルが収集される"""
        # テストファイル作成
        test_file = os.path.join(temp_dir, "test.docx")
        with open(test_file, "w") as f:
            f.write("dummy")

        mock_converter.convert.return_value = "/tmp/converted.pdf"
        mock_processor.get_page_count.return_value = 3

        toc_entries, content_pdfs = collector.collect_documents(temp_dir)

        assert len(content_pdfs) == 1
        mock_converter.convert.assert_called_once()

    def test_cover_file_processed_separately(
        self, collector, temp_dir, mock_converter, mock_processor
    ):
        """表紙ファイルはTOCに含まれずPDFにのみ追加される"""
        # 表紙ファイルとその他のファイル
        cover_file = os.path.join(temp_dir, "表紙.docx")
        other_file = os.path.join(temp_dir, "other.docx")
        with open(cover_file, "w") as f:
            f.write("cover")
        with open(other_file, "w") as f:
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
        all_convert_args = [
            call[0][0] for call in mock_converter.convert.call_args_list
        ]
        assert any("表紙" in arg for arg in all_convert_args)

    def test_cancel_during_collection(self, mock_converter, mock_processor, temp_dir):
        """キャンセルチェックが機能する"""
        collector = DocumentCollector(
            mock_converter, mock_processor, cancel_check=lambda: True
        )

        test_file = os.path.join(temp_dir, "test.docx")
        with open(test_file, "w") as f:
            f.write("dummy")

        with pytest.raises(CancelledError):
            collector.collect_documents(temp_dir)

    def test_directory_with_subdirs(
        self, collector, temp_dir, mock_converter, mock_processor
    ):
        """サブディレクトリが大見出しとして処理される"""
        # サブディレクトリ作成
        sub_dir = os.path.join(temp_dir, "01 教育計画")
        os.makedirs(sub_dir)
        sub_file = os.path.join(sub_dir, "test.docx")
        with open(sub_file, "w") as f:
            f.write("dummy")

        mock_converter.convert.return_value = "/tmp/converted.pdf"
        mock_converter.create_separator_page.return_value = "/tmp/separator.pdf"
        mock_processor.get_page_count.return_value = 1

        toc_entries, content_pdfs = collector.collect_documents(temp_dir)

        # 区切りページが作成される
        mock_converter.create_separator_page.assert_called_once_with("教育計画")
        assert len(content_pdfs) > 0
        # サブディレクトリが大見出しとしてTOCエントリに含まれる
        assert len(toc_entries) == 2
        toc_names = [entry[0] for entry in toc_entries]
        assert "教育計画" in toc_names


@pytest.mark.unit
class TestConvertAndAddPdf:
    """_convert_and_add_pdf のテスト"""

    def test_successful_conversion(
        self, collector, temp_dir, mock_converter, mock_processor
    ):
        """変換成功時にページ数が加算される"""
        test_file = os.path.join(temp_dir, "test.docx")
        with open(test_file, "w") as f:
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
        with open(test_file, "w") as f:
            f.write("dummy")

        mock_converter.convert.return_value = None

        content_pdfs = []
        result = collector._convert_and_add_pdf(test_file, content_pdfs, 10)

        assert result == 10  # 変更なし
        assert len(content_pdfs) == 0


@pytest.mark.unit
class TestProcessRootFile:
    """_process_root_file のテスト"""

    def test_adds_toc_entry(self, collector, temp_dir, mock_converter, mock_processor):
        """TOCエントリが追加される"""
        test_file = os.path.join(temp_dir, "01 概要.docx")
        with open(test_file, "w") as f:
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


@pytest.mark.unit
class TestCoverPageCounting:
    """表紙ページ数カウントのテスト"""

    def test_single_page_cover_does_not_shift_pages(
        self, collector, temp_dir, mock_converter, mock_processor
    ):
        """1ページの表紙はcurrent_pageを変更しない"""
        cover_file = os.path.join(temp_dir, "表紙.docx")
        with open(cover_file, 'w') as f:
            f.write("cover")

        mock_converter.convert.return_value = "/tmp/cover.pdf"
        mock_processor.get_page_count.return_value = 1

        content_pdfs: list = []
        result = collector._process_cover_file(cover_file, content_pdfs, 3)

        # CONTENT_START_PAGE(3)は表紙1ページを前提としているため変化なし
        assert result == 3
        assert len(content_pdfs) == 1
        assert collector.get_cover_pages() == 1

    def test_multi_page_cover_adjusts_offset(
        self, collector, temp_dir, mock_converter, mock_processor
    ):
        """複数ページの表紙は差分のみcurrent_pageに加算"""
        cover_file = os.path.join(temp_dir, "表紙.docx")
        with open(cover_file, 'w') as f:
            f.write("cover")

        mock_converter.convert.return_value = "/tmp/cover.pdf"
        mock_processor.get_page_count.return_value = 3  # 3ページの表紙

        content_pdfs: list = []
        result = collector._process_cover_file(cover_file, content_pdfs, 3)

        # 3ページ表紙: 超過分 = 3 - 1 = 2を加算
        assert result == 5  # 3 + 2
        assert collector.get_cover_pages() == 3

    def test_cover_pages_default(self, mock_converter, mock_processor):
        """表紙未処理時のデフォルトページ数は0（表紙なしを示す）"""
        collector = DocumentCollector(mock_converter, mock_processor)
        assert collector.get_cover_pages() == 0


@pytest.mark.unit
class TestCoverOrdering:
    """表紙ファイルの処理順序のテスト（Unicode順に依存しないこと）"""

    def test_cover_processed_first_even_if_sorted_last(
        self, collector, temp_dir, mock_converter, mock_processor
    ):
        """表紙ファイル名がソートで最後でも、最初に処理されてcontent_pdfsの先頭に置かれる"""
        # 数字フォルダ（ソートで先）と表紙ファイル（Unicode順で後）を配置
        os.makedirs(os.path.join(temp_dir, "01_行事"))
        test_file = os.path.join(temp_dir, "01_行事", "doc.docx")
        with open(test_file, 'w') as f:
            f.write("dummy")
        cover_file = os.path.join(temp_dir, "表紙.docx")
        with open(cover_file, 'w') as f:
            f.write("cover")

        mock_converter.convert.return_value = "/tmp/converted.pdf"
        mock_converter.create_separator_page.return_value = "/tmp/sep.pdf"
        mock_processor.get_page_count.return_value = 1

        toc_entries, content_pdfs = collector.collect_documents(temp_dir)

        # content_pdfsの先頭が表紙
        assert len(content_pdfs) >= 1
        # 表紙ファイル変換結果が先頭に来ているはず
        assert collector.get_cover_pages() == 1
        # 最初のTOCエントリは表紙ページ数+目次後の位置
        # 1ページ表紙 + 1ページ目次 → コンテンツは3ページ目から
        assert toc_entries[0][2] == PDFConstants.CONTENT_START_PAGE

    def test_no_cover_starts_content_at_page_2(
        self, collector, temp_dir, mock_converter, mock_processor
    ):
        """表紙ファイルが無い場合、コンテンツは目次の次の2ページ目から開始"""
        os.makedirs(os.path.join(temp_dir, "01_行事"))
        test_file = os.path.join(temp_dir, "01_行事", "doc.docx")
        with open(test_file, 'w') as f:
            f.write("dummy")

        mock_converter.convert.return_value = "/tmp/converted.pdf"
        mock_converter.create_separator_page.return_value = "/tmp/sep.pdf"
        mock_processor.get_page_count.return_value = 1

        toc_entries, content_pdfs = collector.collect_documents(temp_dir)

        # 表紙なし: _cover_pages = 0
        assert collector.get_cover_pages() == 0
        # 最初のTOCエントリ: 目次(1ページ)の次 = 2
        assert toc_entries[0][2] == 1 + PDFConstants.TOC_PAGE_COUNT


@pytest.mark.unit
class TestSeparatorPageCounting:
    """区切りページの実ページ数計測テスト"""

    def test_separator_uses_measured_page_count(
        self, collector, temp_dir, mock_converter, mock_processor
    ):
        """区切りページの実際のページ数がcurrent_pageに反映される"""
        sub_dir = os.path.join(temp_dir, "01 セクション")
        os.makedirs(sub_dir)
        test_file = os.path.join(sub_dir, "test.docx")
        with open(test_file, 'w') as f:
            f.write("dummy")

        mock_converter.convert.return_value = "/tmp/converted.pdf"
        mock_converter.create_separator_page.return_value = "/tmp/separator.pdf"
        # 区切りページが2ページの場合をシミュレート
        mock_processor.get_page_count.return_value = 2

        toc_entries: list = []
        content_pdfs: list = []
        result = collector._process_subfolder(
            sub_dir, "01 セクション", True,
            toc_entries, content_pdfs, 3
        )

        # 区切り2ページ + ファイル2ページ = 4ページ進む
        assert result == 7  # 3 + 2(separator) + 2(file)
        assert toc_entries[0][2] == 3  # 区切りページの開始ページ
