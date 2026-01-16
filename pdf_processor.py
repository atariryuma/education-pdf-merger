"""
PDF処理モジュール

PDFのマージ、圧縮、ページ番号付加、目次作成などの機能を提供
"""
import logging
import os
import subprocess
from contextlib import contextmanager
from typing import List, Optional, Tuple, TYPE_CHECKING, Generator

import fitz  # PyMuPDF
from PyPDF2 import PdfMerger
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.platypus import BaseDocTemplate, Paragraph, Spacer, PageBreak, Frame, PageTemplate, Table, TableStyle, SimpleDocTemplate
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

from exceptions import PDFProcessingError
from constants import PDFConstants

if TYPE_CHECKING:
    from config_loader import ConfigLoader

# ロガーの設定
logger = logging.getLogger(__name__)


class PDFProcessor:
    """PDF処理を行うクラス"""

    def __init__(self, config: "ConfigLoader") -> None:
        """
        Args:
            config: ConfigLoaderインスタンス
        """
        self.config = config
        self._register_fonts()

    @contextmanager
    def _atomic_pdf_operation(self, pdf_path: str) -> Generator[str, None, None]:
        """
        一時ファイルを使った安全なPDF操作（テンプレートメソッド）

        Args:
            pdf_path: 操作対象のPDFファイルパス

        Yields:
            str: 一時ファイルのパス

        Note:
            処理が成功した場合のみ元ファイルを置換
            失敗時は一時ファイルを自動削除
        """
        tmp_file = pdf_path + PDFConstants.TEMP_FILE_SUFFIX
        tmp_created = False

        try:
            yield tmp_file
            # yieldから正常に戻った = 処理成功
            tmp_created = True
            os.replace(tmp_file, pdf_path)
            tmp_created = False  # replaceで消えたのでクリーンアップ不要
            logger.debug(f"PDF操作完了: {pdf_path}")
        except Exception:
            # 例外は再送出（呼び出し側で処理）
            raise
        finally:
            # エラー時のクリーンアップ（TOCTOU回避）
            if tmp_created:
                try:
                    os.remove(tmp_file)
                    logger.debug(f"一時ファイルを削除: {tmp_file}")
                except FileNotFoundError:
                    pass  # 既に削除済み
                except OSError as e:
                    logger.warning(f"一時ファイル削除失敗: {tmp_file}, エラー: {e}")

    def _register_fonts(self) -> None:
        """フォントを登録"""
        font_path = self.config.get('fonts', 'mincho')
        try:
            pdfmetrics.registerFont(TTFont('Mincho', font_path))
            logger.debug("Minchoフォントを登録しました")
        except Exception as e:
            logger.warning(f"Minchoフォントの登録に失敗しました。フォントファイルを確認してください: {font_path} - {e}")

    def merge_pdfs(self, pdf_paths: List[str], output_file: str) -> None:
        """
        複数のPDFを1つにマージ

        Args:
            pdf_paths: PDFファイルパスのリスト
            output_file: 出力先ファイルパス
        """
        merger = PdfMerger()
        try:
            for pdf in pdf_paths:
                if pdf and os.path.exists(pdf):
                    merger.append(pdf)
            merger.write(output_file)
            logger.info(f"PDFをマージしました: {output_file}")
        finally:
            merger.close()

    def compress_pdf(self, pdf_path: str) -> None:
        """
        GhostscriptでPDFを圧縮.

        Args:
            pdf_path: 圧縮対象のPDFパス
        """
        try:
            with self._atomic_pdf_operation(pdf_path) as tmp_file:
                gs_executable = self.config.get('ghostscript', 'executable')

                gs_command = [
                    gs_executable,
                    "-sDEVICE=pdfwrite",
                    f"-dCompatibilityLevel={PDFConstants.GS_COMPATIBILITY_LEVEL}",
                    f"-dPDFSETTINGS={PDFConstants.GS_PDF_SETTINGS}",
                    "-dNOPAUSE",
                    "-dQUIET",
                    "-dBATCH",
                    f"-sOutputFile={tmp_file}",
                    pdf_path
                ]

                subprocess.run(gs_command, check=True, timeout=PDFConstants.GS_TIMEOUT_SECONDS)
                logger.info(f"Ghostscriptを使用してPDFを圧縮しました: {pdf_path}")
        except subprocess.TimeoutExpired:
            logger.error(f"Ghostscriptがタイムアウトしました: {pdf_path}")
        except subprocess.CalledProcessError as e:
            logger.error(f"Ghostscript実行エラー ({pdf_path}): {e}")
        except Exception as e:
            logger.error(f"PDF圧縮エラー ({pdf_path}): {e}")

    def get_page_count(self, pdf_path: str) -> int:
        """
        PDFのページ数を取得

        Args:
            pdf_path: PDFファイルパス

        Returns:
            int: ページ数

        Raises:
            PDFProcessingError: ページ数の取得に失敗した場合
        """
        try:
            with fitz.open(pdf_path) as doc:
                page_count = doc.page_count

                # 破損したPDFの場合は警告をログ出力
                if doc.is_repaired:
                    logger.warning(f"PDFが修復されました: {pdf_path}")

                return page_count
        except Exception as e:
            logger.error(f"ページ数の取得に失敗しました: {pdf_path} - {e}")
            raise PDFProcessingError(
                f"PDFファイルの読み込みに失敗: {pdf_path}",
                operation="ページ数取得",
                original_error=e
            ) from e

    def add_page_numbers(self, pdf_file: str, exclude_first_pages: int = PDFConstants.COVER_PAGE_COUNT) -> None:
        """
        PDFにページ番号を追加

        Args:
            pdf_file: PDFファイルパス
            exclude_first_pages: ページ番号を表示しない先頭ページ数
        """
        with self._atomic_pdf_operation(pdf_file) as tmp_file:
            with fitz.open(pdf_file) as doc:
                total_pages = doc.page_count

                for i in range(total_pages):
                    # 先頭ページはスキップ
                    if i < exclude_first_pages:
                        continue

                    number_text = str(i + 1)
                    page = doc.load_page(i)
                    rect = page.rect
                    # ページ中央下部に配置
                    point = fitz.Point(
                        rect.width / 2 - PDFConstants.PAGE_NUMBER_X_OFFSET,
                        rect.height - PDFConstants.PAGE_NUMBER_BOTTOM_MARGIN
                    )
                    page.insert_text(
                        point, number_text,
                        fontsize=PDFConstants.PAGE_NUMBER_FONT_SIZE,
                        fontname=PDFConstants.PAGE_NUMBER_FONT_NAME,
                        color=(0, 0, 0)
                    )

                doc.save(tmp_file)

            logger.info(f"ページ番号を追加しました: {pdf_file} (先頭{exclude_first_pages}ページはスキップ)")

    def set_pdf_outlines(self, pdf_file: str, toc_entries: List[Tuple[str, int, int]]) -> None:
        """
        PDFにアウトライン（しおり）を設定

        Args:
            pdf_file: PDFファイルパス
            toc_entries: 目次エントリのリスト [(title, level, page), ...]
        """
        with self._atomic_pdf_operation(pdf_file) as tmp_file:
            with fitz.open(pdf_file) as doc:
                page_count = doc.page_count

                corrected_outlines = []
                for title, level, page in toc_entries:
                    # ページ番号を有効範囲に補正
                    if page < 1:
                        page = 1
                    if page > page_count:
                        page = page_count
                    corrected_outlines.append([level, title, page])

                # PyMuPDFの制約：最初の項目は必ずレベル1
                if corrected_outlines and corrected_outlines[0][0] != PDFConstants.HEADING_LEVEL_MAIN:
                    corrected_outlines[0][0] = PDFConstants.HEADING_LEVEL_MAIN

                logger.debug(f"PDFアウトラインを設定: {corrected_outlines}")

                try:
                    doc.set_toc(corrected_outlines)
                except Exception as e:
                    logger.error(f"PDFアウトラインの設定に失敗しました: {e}")

                doc.save(tmp_file, incremental=False)

            logger.info("PDFアウトライン（しおり）を設定しました")

    def create_toc_pdf(self, toc_entries: List[Tuple[str, int, int]], output_path: str) -> str:
        """
        目次ページのPDFを作成

        Args:
            toc_entries: 目次エントリのリスト [(title, level, page), ...]
            output_path: 出力先PDFパス

        Returns:
            str: 作成したPDFのパス
        """
        doc = BaseDocTemplate(output_path, pagesize=A4)
        frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height)
        doc.addPageTemplates([PageTemplate(id='normal', frames=[frame])])

        # テーブルデータの作成
        table_data = []
        for title, level, page in toc_entries:
            # レベル2の項目はインデント
            indent = "    " if level == 2 else ""
            title_text = indent + title
            table_data.append([title_text, str(page)])

        if not table_data:
            table_data = [["目次なし", ""]]

        # テーブルの作成
        col_widths = [doc.width * 0.8, doc.width * 0.2]
        toc_table = Table(table_data, colWidths=col_widths)
        toc_table.setStyle(TableStyle([
            ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('FONTNAME', (0, 0), (-1, -1), 'Mincho'),
            ('FONTSIZE', (0, 0), (-1, -1), 12),
            ('LEFTPADDING', (0, 0), (-1, -1), 2),
            ('RIGHTPADDING', (0, 0), (-1, -1), 2),
            ('LINEBELOW', (0, 0), (-1, -1), 0.25, colors.grey),
        ]))

        # ストーリーの構築
        story = []
        title_style = ParagraphStyle('toc_title', fontName='Mincho', fontSize=18, spaceAfter=12)
        story.append(Paragraph("目次", title_style))
        story.append(Spacer(1, 0.2 * inch))
        story.append(toc_table)
        story.append(PageBreak())

        doc.build(story)
        logger.info(f"目次PDFを作成しました: {output_path}")
        return output_path

    def create_separator_pdf(self, title: str, output_path: str) -> str:
        """
        区切りページのPDFを作成（reportlab完全生成）

        Args:
            title: セクションタイトル
            output_path: 出力先PDFパス

        Returns:
            str: 作成したPDFのパス
        """
        # SimpleDocTemplateで1ページ生成
        doc = SimpleDocTemplate(output_path, pagesize=A4)

        # タイトルスタイル: 明朝体24pt、中央揃え
        title_style = ParagraphStyle(
            'separator_title',
            fontName='Mincho',
            fontSize=24,
            alignment=TA_CENTER,
            leading=36  # 行間
        )

        # ストーリー構築: 縦中央配置
        story = []
        story.append(Spacer(1, 3.5 * inch))  # 上部スペース（A4の約1/3）
        story.append(Paragraph(title, title_style))
        story.append(Spacer(1, 3.5 * inch))  # 下部スペース

        # PDF生成
        doc.build(story)
        logger.info(f"区切りページを生成: {title}")
        return output_path

    def split_pdf(self, pdf_path: str, output_dir: str) -> Tuple[str, str]:
        """
        PDFを表紙と残りのページに分割

        Args:
            pdf_path: 分割対象のPDFパス
            output_dir: 出力先ディレクトリ

        Returns:
            tuple: (表紙PDFパス, 残りのPDFパス)

        Raises:
            PDFProcessingError: 分割処理に失敗した場合
        """
        try:
            cover_pdf = os.path.join(output_dir, "cover.pdf")
            remainder_pdf = os.path.join(output_dir, "remainder.pdf")

            with fitz.open(pdf_path) as doc:
                # 表紙（1ページ目）
                with fitz.open() as cover_doc:
                    cover_doc.insert_pdf(doc, from_page=0, to_page=0)
                    cover_doc.save(cover_pdf)

                # 残りのページ
                with fitz.open() as remainder_doc:
                    if doc.page_count > 1:
                        remainder_doc.insert_pdf(doc, from_page=1, to_page=doc.page_count - 1)
                    remainder_doc.save(remainder_pdf)

            logger.debug(f"PDFを分割しました: 表紙={cover_pdf}, 残り={remainder_pdf}")
            return cover_pdf, remainder_pdf

        except Exception as e:
            logger.error(f"PDF分割エラー ({pdf_path}): {e}")
            raise PDFProcessingError(
                f"PDFの分割に失敗: {pdf_path}",
                operation="分割",
                original_error=e
            ) from e
