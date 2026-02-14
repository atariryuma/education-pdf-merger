"""
PDF結合オーケストレーター

全体の処理フローを制御し、各コンポーネントを協調させる
"""
import logging
import os
from typing import Callable, Optional

from config_loader import ConfigLoader
from pdf_converter import PDFConverter
from pdf_processor import PDFProcessor
from document_collector import DocumentCollector
from exceptions import CancelledError
from constants import PDFConstants

logger = logging.getLogger(__name__)


class PDFMergeOrchestrator:
    """PDF結合のオーケストレーター（全体の流れを制御）"""

    def __init__(
        self,
        config: ConfigLoader,
        pdf_converter: PDFConverter,
        pdf_processor: PDFProcessor,
        document_collector: DocumentCollector,
        cancel_check: Optional[Callable[[], bool]] = None
    ) -> None:
        """
        Args:
            config: ConfigLoaderインスタンス
            pdf_converter: PDFConverterインスタンス
            pdf_processor: PDFProcessorインスタンス
            document_collector: DocumentCollectorインスタンス
            cancel_check: キャンセル状態をチェックするコールバック関数
        """
        self.config = config
        self.converter = pdf_converter
        self.processor = pdf_processor
        self.collector = document_collector
        self.temp_dir = config.get_temp_dir()
        self._cancel_check = cancel_check or (lambda: False)

    def is_cancelled(self) -> bool:
        """キャンセルされたかどうかを確認"""
        return self._cancel_check()

    def _check_cancel(self) -> None:
        """キャンセルされていれば例外を投げる"""
        if self.is_cancelled():
            raise CancelledError("処理がキャンセルされました")

    def _cleanup_temp_files(self, *file_paths: str) -> None:
        """
        一時ファイルを安全に削除

        Args:
            *file_paths: 削除対象のファイルパス
        """
        for file_path in file_paths:
            try:
                if file_path and os.path.exists(file_path):
                    os.remove(file_path)
                    logger.debug(f"一時ファイルを削除: {file_path}")
            except OSError as e:
                logger.warning(f"一時ファイル削除失敗: {file_path} - {e}")

    def create_merged_pdf(
        self,
        target_dir: str,
        output_pdf: str,
        create_separator_for_subfolder: bool = True
    ) -> None:
        """
        ディレクトリ内のドキュメントを統合したPDFを作成

        Args:
            target_dir: 処理対象のディレクトリ
            output_pdf: 出力PDFのパス
            create_separator_for_subfolder: サブフォルダに区切りページを作成するか

        Raises:
            CancelledError: 処理がキャンセルされた場合
            PDFProcessingError: PDF処理中にエラーが発生した場合
        """
        logger.info(f"PDFマージ処理を開始: {target_dir}")
        logger.info(f"出力先: {output_pdf}")

        # 一時ファイルのパスを定義（finallyでクリーンアップ用）
        temp_merged = os.path.join(self.temp_dir, "temp_merged.pdf")
        toc_pdf = os.path.join(self.temp_dir, "toc.pdf")
        cover_pdf = ""
        remainder_pdf = ""

        try:
            # 1. ドキュメント収集とPDF変換
            logger.info("[Step 1/6] ドキュメントを収集・変換中...")
            toc_entries, content_pdfs = self.collector.collect_documents(
                target_dir,
                create_separator_for_subfolder
            )
            self._check_cancel()

            # 2. 一時的にマージ
            logger.info("[Step 2/6] 一時マージPDFを作成中...")
            self.processor.merge_pdfs(content_pdfs, temp_merged)
            self._check_cancel()

            # 3. 目次PDFを生成
            logger.info("[Step 3/6] 目次を作成中...")
            self.processor.create_toc_pdf(toc_entries, toc_pdf)
            self._check_cancel()

            # 4. 表紙と残りのページに分割
            logger.info("[Step 4/6] 表紙とコンテンツを分割中...")
            cover_pdf, remainder_pdf = self.processor.split_pdf(temp_merged, self.temp_dir)
            self._check_cancel()

            # 5. 最終的にマージ（表紙 + 目次 + 残り）
            logger.info("[Step 5/6] 最終PDFをマージ中...")
            final_list = [cover_pdf, toc_pdf, remainder_pdf]
            self.processor.merge_pdfs(final_list, output_pdf)
            self._check_cancel()

            # 6. ページ番号を追加（表紙は除外）
            logger.info("[Step 6/6] ページ番号としおりを追加中...")
            self.processor.add_page_numbers(output_pdf, exclude_first_pages=PDFConstants.COVER_PAGE_COUNT)
            self._check_cancel()

            # 7. PDFアウトライン（しおり）を設定
            self.processor.set_pdf_outlines(output_pdf, toc_entries)

            total_pages = self.processor.get_page_count(output_pdf)
            logger.info(f"PDFの作成が完了しました: {output_pdf}")
            logger.info(f"  目次エントリ数: {len(toc_entries)}")
            logger.info(f"  総ページ数: {total_pages}")
        finally:
            # 一時ファイルをクリーンアップ
            self._cleanup_temp_files(temp_merged, toc_pdf, cover_pdf, remainder_pdf)
