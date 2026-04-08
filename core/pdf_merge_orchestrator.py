"""
PDF結合オーケストレーター

全体の処理フローを制御し、各コンポーネントを協調させる
"""
import logging
import os
import uuid
from typing import Callable, List, Optional, Tuple

from infrastructure.config_loader import ConfigLoader
from core.pdf_converter import PDFConverter
from core.pdf_processor import PDFProcessor
from core.document_collector import DocumentCollector
from shared.exceptions import CancelledError
from shared.constants import PDFConstants
from infrastructure.ghostscript import GhostscriptDetector

logger = logging.getLogger(__name__)


class PDFMergeOrchestrator:
    """PDF結合のオーケストレーター（全体の流れを制御）"""

    def __init__(
        self,
        config: ConfigLoader,
        pdf_converter: PDFConverter,
        pdf_processor: PDFProcessor,
        document_collector: DocumentCollector,
        cancel_check: Optional[Callable[[], bool]] = None,
        progress_callback: Optional[Callable[[int, int, str], None]] = None,
    ) -> None:
        """
        Args:
            config: ConfigLoaderインスタンス
            pdf_converter: PDFConverterインスタンス
            pdf_processor: PDFProcessorインスタンス
            document_collector: DocumentCollectorインスタンス
            cancel_check: キャンセル状態をチェックするコールバック関数
            progress_callback: 進捗通知コールバック (step, total, message)
        """
        self.config = config
        self.converter = pdf_converter
        self.processor = pdf_processor
        self.collector = document_collector
        self.temp_dir = config.get_temp_dir()
        normalized = os.path.normcase(os.path.abspath(self.temp_dir))
        self._normalized_temp_dir = normalized
        self._normalized_temp_prefix = normalized + os.sep
        self._cancel_check = cancel_check or (lambda: False)
        self._user_progress_callback = progress_callback or (
            lambda step, total, msg: None
        )

    def is_cancelled(self) -> bool:
        """キャンセルされたかどうかを確認"""
        return self._cancel_check()

    def _progress_callback(self, step: int, total: int, message: str) -> None:
        """進捗をログとコールバックの両方に通知"""
        logger.info(f"[Step {step}/{total}] {message}")
        self._user_progress_callback(step, total, message)

    def _check_cancel(self) -> None:
        """キャンセルされていれば例外を投げる"""
        if self.is_cancelled():
            raise CancelledError("処理がキャンセルされました")

    def _is_temp_file(self, file_path: str) -> bool:
        """一時ディレクトリ内のファイルかどうかを判定"""
        try:
            normalized = os.path.normcase(os.path.abspath(file_path))
            return (
                normalized.startswith(self._normalized_temp_prefix)
                or normalized == self._normalized_temp_dir
            )
        except (ValueError, OSError):
            return False

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
        toc_entries: List[Tuple[str, int, int]], page_offset: int
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

    @staticmethod
    def _validate_toc_entries(
        toc_entries: List[Tuple[str, int, int]], total_pages: int
    ) -> None:
        """
        目次エントリのページ番号が最終PDFの有効範囲内か検証する

        Args:
            toc_entries: 目次エントリ [(title, level, page), ...]
            total_pages: 最終PDFの総ページ数
        """
        for title, _level, page in toc_entries:
            if page < 1 or page > total_pages:
                logger.warning(
                    "目次エントリ '%s' のページ番号 %d が有効範囲外です (1-%d)",
                    title, page, total_pages
                )

    def _compress_output(self, output_pdf: str) -> None:
        """
        Ghostscriptで最終PDFを圧縮

        Ghostscriptが利用不可の場合は警告ログを出してスキップする。

        Args:
            output_pdf: 圧縮対象のPDFパス
        """
        gs_path = self.config.get("ghostscript", "executable")
        if not gs_path:
            gs_path = GhostscriptDetector.detect()
            if gs_path:
                # compress_pdf() は config から GS パスを読むため、検出結果を保存する
                self.config.set("ghostscript", "executable", value=gs_path)

        if not gs_path or not GhostscriptDetector.validate_ghostscript(gs_path):
            logger.warning("Ghostscriptが見つからないため、圧縮をスキップします")
            return

        original_size = os.path.getsize(output_pdf)
        success = self.processor.compress_pdf(output_pdf)

        if success:
            compressed_size = os.path.getsize(output_pdf)
            ratio = (
                (1 - compressed_size / original_size) * 100 if original_size > 0 else 0
            )
            logger.info(
                "PDF圧縮完了: %.1fMB → %.1fMB (%.0f%%削減)",
                original_size / 1_048_576,
                compressed_size / 1_048_576,
                ratio,
            )
        else:
            logger.warning("PDF圧縮に失敗しました。圧縮前のファイルを使用します")

    _MAX_TOC_ITERATIONS = 5

    def _create_stable_toc_pdf(
        self, base_toc_entries: List[Tuple[str, int, int]], toc_pdf: str
    ) -> List[Tuple[str, int, int]]:
        """
        実際の目次ページ数に合わせて、目次エントリを補正しながら目次PDFを作成する

        Note:
            DocumentCollector は「表紙N + 目次1」前提でページ番号を採番するため、
            目次が複数ページにわたる場合は補正が必要。
        """
        assumed_toc_pages = PDFConstants.TOC_PAGE_COUNT
        current_toc_pages = assumed_toc_pages
        adjusted_toc_entries = list(base_toc_entries)

        # 目次ページ数が変化しなくなるまで再計算
        for iteration in range(self._MAX_TOC_ITERATIONS):
            page_offset = current_toc_pages - assumed_toc_pages
            adjusted_toc_entries = self._offset_toc_entries(
                base_toc_entries, page_offset
            )

            measured_toc_pages = self.processor.create_toc_pdf(
                adjusted_toc_entries, toc_pdf
            )

            if measured_toc_pages == current_toc_pages:
                if page_offset != 0:
                    logger.info(
                        "目次ページ数補正: assumed=%s, actual=%s, offset=%s (収束: %d回)",
                        assumed_toc_pages,
                        measured_toc_pages,
                        page_offset,
                        iteration + 1
                    )
                return adjusted_toc_entries

            current_toc_pages = measured_toc_pages

        # 念のため最終値で再生成して返す
        page_offset = current_toc_pages - assumed_toc_pages
        adjusted_toc_entries = self._offset_toc_entries(base_toc_entries, page_offset)
        self.processor.create_toc_pdf(adjusted_toc_entries, toc_pdf)
        logger.warning(
            "目次ページ数補正が収束しませんでした。最後の計算結果を採用します: actual=%s",
            current_toc_pages,
        )
        return adjusted_toc_entries

    def create_merged_pdf(
        self,
        target_dir: str,
        output_pdf: str,
        create_separator_for_subfolder: bool = True,
        compress: bool = False,
    ) -> None:
        """
        ディレクトリ内のドキュメントを統合したPDFを作成

        Args:
            target_dir: 処理対象のディレクトリ
            output_pdf: 出力PDFのパス
            create_separator_for_subfolder: サブフォルダに区切りページを作成するか
            compress: Ghostscriptで最終PDFを圧縮するか

        Raises:
            CancelledError: 処理がキャンセルされた場合
            PDFProcessingError: PDF処理中にエラーが発生した場合
        """
        logger.info(f"PDFマージ処理を開始: {target_dir}")
        logger.info(f"出力先: {output_pdf}")

        total_steps = (
            PDFConstants.MERGE_STEPS_WITH_COMPRESS
            if compress
            else PDFConstants.MERGE_STEPS
        )

        # 一時ファイルのパスを定義（finallyでクリーンアップ用）
        unique_id = uuid.uuid4().hex[:8]
        temp_merged = os.path.join(self.temp_dir, f"temp_merged_{unique_id}.pdf")
        toc_pdf = os.path.join(self.temp_dir, f"toc_{unique_id}.pdf")
        cover_pdf: Optional[str] = None
        remainder_pdf: Optional[str] = None
        content_pdfs: List[str] = []

        try:
            # 1. ドキュメント収集とPDF変換
            self._progress_callback(1, total_steps, "ドキュメントを収集・変換中...")
            toc_entries, content_pdfs = self.collector.collect_documents(
                target_dir, create_separator_for_subfolder
            )
            self._check_cancel()

            # 2. 一時的にマージ
            self._progress_callback(2, total_steps, "一時マージPDFを作成中...")
            self.processor.merge_pdfs(content_pdfs, temp_merged)
            self._check_cancel()

            # 3. 目次PDFを生成
            self._progress_callback(3, total_steps, "目次を作成中...")
            adjusted_toc_entries = self._create_stable_toc_pdf(toc_entries, toc_pdf)
            self._check_cancel()

            # 4. 表紙と残りのページに分割（実際の表紙ページ数を使用）
            self._progress_callback(4, total_steps, "表紙とコンテンツを分割中...")
            cover_pages = self.collector.get_cover_pages()
            cover_pdf, remainder_pdf = self.processor.split_pdf(
                temp_merged, self.temp_dir, cover_pages
            )
            self._check_cancel()

            # 5. 最終的にマージ（表紙 + 目次 + 残り）
            self._progress_callback(5, total_steps, "最終PDFをマージ中...")
            final_list = [
                p for p in [cover_pdf, toc_pdf, remainder_pdf] if p is not None
            ]
            self.processor.merge_pdfs(final_list, output_pdf)
            self._check_cancel()

            # 6. ページ番号を追加（表紙は除外）・しおり設定
            self._progress_callback(6, total_steps, "ページ番号としおりを追加中...")
            self.processor.add_page_numbers(
                output_pdf, exclude_first_pages=cover_pages
            )
            self.processor.set_pdf_outlines(output_pdf, adjusted_toc_entries)
            self._check_cancel()

            # 7. Ghostscript圧縮（オプション）
            if compress:
                self._progress_callback(7, total_steps, "PDFを圧縮中...")
                self._compress_output(output_pdf)

            # 最終検証: 目次エントリのページ番号が有効範囲内か確認
            total_pages = self.processor.get_page_count(output_pdf)
            self._validate_toc_entries(adjusted_toc_entries, total_pages)

            logger.info(f"PDFの作成が完了しました: {output_pdf}")
            logger.info(f"  目次エントリ数: {len(adjusted_toc_entries)}")
            logger.info(f"  総ページ数: {total_pages}")
        finally:
            # content_pdfsにはソースPDFの元パスが含まれうる（PDFConverter.convert()は
            # .pdfファイルを変換せず元パスをそのまま返すため）。
            # _is_temp_fileフィルタにより、一時ディレクトリ内のファイルのみ削除する。
            all_content_pdfs = content_pdfs or self.collector.get_collected_pdfs()
            files_to_cleanup = [
                p
                for p in [temp_merged, toc_pdf, cover_pdf, remainder_pdf]
                if p is not None
            ]
            files_to_cleanup.extend(
                p for p in all_content_pdfs if self._is_temp_file(p)
            )
            self._cleanup_temp_files(*files_to_cleanup)
