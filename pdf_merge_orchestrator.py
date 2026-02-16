"""
PDF結合オーケストレーター

全体の処理フローを制御し、各コンポーネントを協調させる
"""
import logging
import os
from typing import Callable, List, Optional, Tuple

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

    @staticmethod
    def _offset_toc_entries(
        toc_entries: List[Tuple[str, int, int]],
        page_offset: int
    ) -> List[Tuple[str, int, int]]:
        """
        目次エントリのページ番号にオフセットを加算する

        Args:
            toc_entries: 目次エントリ [(title, level, page), ...]
            page_offset: 加算するページ数

        Returns:
            補正後の目次エントリ
        """
        if page_offset == 0:
            return list(toc_entries)

        adjusted_entries: List[Tuple[str, int, int]] = []
        for title, level, page in toc_entries:
            adjusted_entries.append((title, level, max(1, page + page_offset)))
        return adjusted_entries

    def _create_stable_toc_pdf(
        self,
        base_toc_entries: List[Tuple[str, int, int]],
        toc_pdf: str
    ) -> List[Tuple[str, int, int]]:
        """
        実際の目次ページ数に合わせて、目次エントリを補正しながら目次PDFを作成する

        Note:
            DocumentCollector は「表紙1 + 目次1」前提でページ番号を採番するため、
            目次が複数ページにわたる場合は補正が必要。
        """
        assumed_toc_pages = PDFConstants.TOC_PAGE_COUNT
        current_toc_pages = assumed_toc_pages
        adjusted_toc_entries = list(base_toc_entries)

        # 目次ページ数が変化しなくなるまで再計算
        for _ in range(3):
            page_offset = current_toc_pages - assumed_toc_pages
            adjusted_toc_entries = self._offset_toc_entries(base_toc_entries, page_offset)

            self.processor.create_toc_pdf(adjusted_toc_entries, toc_pdf)
            measured_toc_pages = self.processor.get_page_count(toc_pdf)

            if measured_toc_pages == current_toc_pages:
                if page_offset != 0:
                    logger.info(
                        "目次ページ数補正: assumed=%s, actual=%s, offset=%s",
                        assumed_toc_pages,
                        measured_toc_pages,
                        page_offset
                    )
                return adjusted_toc_entries

            current_toc_pages = measured_toc_pages

        # 念のため最終値で再生成して返す
        page_offset = current_toc_pages - assumed_toc_pages
        adjusted_toc_entries = self._offset_toc_entries(base_toc_entries, page_offset)
        self.processor.create_toc_pdf(adjusted_toc_entries, toc_pdf)
        logger.warning(
            "目次ページ数補正が収束しませんでした。最後の計算結果を採用します: actual=%s",
            current_toc_pages
        )
        return adjusted_toc_entries

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
            adjusted_toc_entries = self._create_stable_toc_pdf(toc_entries, toc_pdf)
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
            self.processor.set_pdf_outlines(output_pdf, adjusted_toc_entries)

            total_pages = self.processor.get_page_count(output_pdf)
            logger.info(f"PDFの作成が完了しました: {output_pdf}")
            logger.info(f"  目次エントリ数: {len(adjusted_toc_entries)}")
            logger.info(f"  総ページ数: {total_pages}")
        finally:
            # 一時ファイルをクリーンアップ
            self._cleanup_temp_files(temp_merged, toc_pdf, cover_pdf, remainder_pdf)
