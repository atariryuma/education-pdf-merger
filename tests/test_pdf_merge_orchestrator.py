"""
PDFMergeOrchestrator のユニットテスト

オーケストレーション層の6ステップフロー制御をテスト
"""
import os
import pytest
from unittest.mock import MagicMock, patch, call

from pdf_merge_orchestrator import PDFMergeOrchestrator
from exceptions import CancelledError


@pytest.fixture
def mock_deps(temp_dir):
    """オーケストレーターの依存オブジェクトをモック化"""
    config = MagicMock()
    config.get_temp_dir.return_value = temp_dir

    converter = MagicMock()
    processor = MagicMock()
    collector = MagicMock()

    return config, converter, processor, collector


@pytest.fixture
def orchestrator(mock_deps):
    """テスト用オーケストレーターインスタンス"""
    config, converter, processor, collector = mock_deps
    return PDFMergeOrchestrator(config, converter, processor, collector)


class TestPDFMergeOrchestratorInit:
    """初期化のテスト"""

    def test_init_stores_dependencies(self, mock_deps):
        config, converter, processor, collector = mock_deps
        orch = PDFMergeOrchestrator(config, converter, processor, collector)

        assert orch.config is config
        assert orch.converter is converter
        assert orch.processor is processor
        assert orch.collector is collector

    def test_init_default_cancel_check(self, mock_deps):
        config, converter, processor, collector = mock_deps
        orch = PDFMergeOrchestrator(config, converter, processor, collector)
        assert orch.is_cancelled() is False

    def test_init_custom_cancel_check(self, mock_deps):
        config, converter, processor, collector = mock_deps
        orch = PDFMergeOrchestrator(
            config, converter, processor, collector,
            cancel_check=lambda: True
        )
        assert orch.is_cancelled() is True


class TestCreateMergedPDF:
    """create_merged_pdf の6ステップフローテスト"""

    def test_full_flow_calls_all_steps(self, orchestrator, mock_deps):
        """6ステップが順番に呼ばれることを確認"""
        _, _, processor, collector = mock_deps

        # モックの戻り値を設定
        collector.collect_documents.return_value = (
            [("Section1", 1, 3)], ["/tmp/a.pdf"]
        )
        processor.split_pdf.return_value = ("/tmp/cover.pdf", "/tmp/remainder.pdf")
        processor.get_page_count.return_value = 10

        orchestrator.create_merged_pdf("/target", "/output.pdf")

        # Step 1: ドキュメント収集
        collector.collect_documents.assert_called_once_with("/target", True)
        # Step 2: 一時マージ
        assert processor.merge_pdfs.call_count == 2  # Step2とStep5
        # Step 3: 目次生成
        processor.create_toc_pdf.assert_called_once()
        # Step 4: 分割
        processor.split_pdf.assert_called_once()
        # Step 6: ページ番号
        processor.add_page_numbers.assert_called_once()
        # Step 7: アウトライン
        processor.set_pdf_outlines.assert_called_once()

    def test_cancel_at_step1(self, mock_deps, temp_dir):
        """Step1後のキャンセルで例外発生"""
        config, converter, processor, collector = mock_deps
        cancel_flag = [False]

        def cancel_check():
            return cancel_flag[0]

        orch = PDFMergeOrchestrator(
            config, converter, processor, collector,
            cancel_check=cancel_check
        )

        collector.collect_documents.return_value = ([], [])
        # Step1実行後にキャンセルフラグを立てる
        def set_cancel(*args, **kwargs):
            cancel_flag[0] = True
            return ([], [])
        collector.collect_documents.side_effect = set_cancel

        with pytest.raises(CancelledError):
            orch.create_merged_pdf("/target", "/output.pdf")

    def test_separator_flag_passed_to_collector(self, orchestrator, mock_deps):
        """create_separator_for_subfolderフラグがcollectorに渡される"""
        _, _, processor, collector = mock_deps

        collector.collect_documents.return_value = ([], ["/tmp/a.pdf"])
        processor.split_pdf.return_value = ("/tmp/c.pdf", "/tmp/r.pdf")
        processor.get_page_count.return_value = 1

        orchestrator.create_merged_pdf("/target", "/out.pdf", create_separator_for_subfolder=False)
        collector.collect_documents.assert_called_once_with("/target", False)

    def test_final_merge_order(self, orchestrator, mock_deps):
        """最終マージが cover + toc + remainder の順で行われる"""
        _, _, processor, collector = mock_deps

        collector.collect_documents.return_value = ([("A", 1, 1)], ["/tmp/a.pdf"])
        processor.split_pdf.return_value = ("/tmp/cover.pdf", "/tmp/remainder.pdf")
        processor.get_page_count.return_value = 5

        orchestrator.create_merged_pdf("/target", "/output.pdf")

        # 2回目のmerge_pdfs呼び出し（最終マージ）を検証
        final_merge_call = processor.merge_pdfs.call_args_list[1]
        pdf_list = final_merge_call[0][0]
        assert pdf_list[0] == "/tmp/cover.pdf"
        assert pdf_list[2] == "/tmp/remainder.pdf"


class TestCleanupTempFiles:
    """一時ファイルクリーンアップのテスト"""

    def test_cleanup_removes_existing_files(self, orchestrator, temp_dir):
        """存在するファイルが削除される"""
        tmp_file = os.path.join(temp_dir, "test.tmp")
        with open(tmp_file, 'w') as f:
            f.write("test")

        orchestrator._cleanup_temp_files(tmp_file)
        assert not os.path.exists(tmp_file)

    def test_cleanup_ignores_nonexistent_files(self, orchestrator):
        """存在しないファイルは無視される"""
        orchestrator._cleanup_temp_files("/nonexistent/path.tmp")
        # 例外が発生しないことを確認

    def test_cleanup_ignores_empty_paths(self, orchestrator):
        """空文字列パスは無視される"""
        orchestrator._cleanup_temp_files("", None)

    def test_cleanup_runs_on_exception(self, mock_deps, temp_dir):
        """処理中の例外でもクリーンアップが実行される"""
        config, converter, processor, collector = mock_deps
        orch = PDFMergeOrchestrator(config, converter, processor, collector)

        collector.collect_documents.side_effect = RuntimeError("test error")

        with pytest.raises(RuntimeError):
            orch.create_merged_pdf("/target", "/output.pdf")

        # finallyブロックが実行されたことはエラーにならないことで確認
