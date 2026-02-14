"""
ドキュメント収集モジュール

ディレクトリ構造を探索し、ファイルをPDF化して目次を作成
"""
import logging
import os
from typing import List, Tuple, Callable, Optional

from pdf_converter import PDFConverter
from pdf_processor import PDFProcessor
from exceptions import CancelledError
from constants import PDFConstants

# ロガーの設定
logger = logging.getLogger(__name__)


class DocumentCollector:
    """
    ディレクトリを探索してドキュメントを収集し、目次を生成するクラス

    定数はPDFConstantsクラスから参照します。
    """

    def __init__(
        self,
        pdf_converter: PDFConverter,
        pdf_processor: PDFProcessor,
        cancel_check: Optional[Callable[[], bool]] = None
    ) -> None:
        """
        Args:
            pdf_converter: PDFConverterインスタンス
            pdf_processor: PDFProcessorインスタンス
            cancel_check: キャンセル状態をチェックするコールバック関数
        """
        self.converter = pdf_converter
        self.processor = pdf_processor
        self._cancel_check = cancel_check or (lambda: False)

    def is_cancelled(self) -> bool:
        """キャンセルされたかどうかを確認"""
        return self._cancel_check()

    @staticmethod
    def _sanitize_name(name: str) -> str:
        """名前から先頭の数字とスペース、アンダースコアを除去"""
        return name.lstrip("0123456789 ").strip().replace("_", "")

    def _convert_and_add_pdf(
        self,
        file_path: str,
        content_pdfs: List[str],
        current_page: int
    ) -> int:
        """
        ファイルをPDFに変換してリストに追加し、更新されたページ番号を返す

        Args:
            file_path: 変換対象ファイルのパス
            content_pdfs: PDFパスのリスト（破壊的に更新される）
            current_page: 現在のページ番号

        Returns:
            int: 更新後のページ番号（変換成功時はページ数が加算される）

        Note:
            変換失敗時はページ番号は変更されず、警告ログが出力される
        """
        file_name = os.path.basename(file_path)
        logger.info(f"ファイルを変換中: {file_name}")
        converted_pdf = self.converter.convert(file_path)
        if converted_pdf:
            page_count = self.processor.get_page_count(converted_pdf)
            content_pdfs.append(converted_pdf)
            logger.info(f"変換成功: {file_name} ({page_count}ページ)")
            return current_page + page_count
        else:
            logger.warning(f"変換スキップ: {file_name}")
        return current_page

    def _process_cover_file(
        self,
        file_path: str,
        content_pdfs: List[str],
        current_page: int
    ) -> int:
        """
        表紙ファイルを処理

        Args:
            file_path: 表紙ファイルのパス
            content_pdfs: PDFパスのリスト（破壊的に更新される）
            current_page: 現在のページ番号

        Returns:
            int: 更新後のページ番号
        """
        logger.info(f"表紙ファイルを処理中: {os.path.basename(file_path)}")
        return self._convert_and_add_pdf(file_path, content_pdfs, current_page)

    def _process_subfolder(
        self,
        subfolder_path: str,
        subfolder_name: str,
        create_separator: bool,
        toc_entries: List[Tuple[str, int, int]],
        content_pdfs: List[str],
        current_page: int
    ) -> int:
        """
        サブフォルダを処理

        Args:
            subfolder_path: サブフォルダのパス
            subfolder_name: サブフォルダ名
            create_separator: 区切りページを作成するか
            toc_entries: 目次エントリのリスト（破壊的に更新される）
            content_pdfs: PDFパスのリスト（破壊的に更新される）
            current_page: 現在のページ番号

        Returns:
            int: 更新後のページ番号
        """
        sub_heading = self._sanitize_name(subfolder_name)
        logger.info(f"サブフォルダを処理中: {subfolder_name}")

        if create_separator:
            # サブフォルダにも区切りページを作成
            logger.info(f"区切りページを作成: {sub_heading}")
            sep_pdf = self.converter.create_separator_page(sub_heading)
            if sep_pdf:
                content_pdfs.append(sep_pdf)
                toc_entries.append((sub_heading, PDFConstants.HEADING_LEVEL_SUB, current_page))
                current_page += 1
        else:
            # 区切りページなしで目次にのみ登録
            toc_entries.append((sub_heading, PDFConstants.HEADING_LEVEL_SUB, current_page))

        # サブフォルダ内のファイルを処理（ディレクトリは除外）
        all_items = sorted(os.listdir(subfolder_path))
        files = [f for f in all_items if os.path.isfile(os.path.join(subfolder_path, f))]
        logger.info(f"サブフォルダ内のファイル数: {len(files)} (全アイテム: {len(all_items)})")
        for filename in files:
            file_path = os.path.join(subfolder_path, filename)
            current_page = self._convert_and_add_pdf(file_path, content_pdfs, current_page)

        return current_page

    def _process_directory(
        self,
        dir_path: str,
        dir_name: str,
        create_separator_for_subfolder: bool,
        toc_entries: List[Tuple[str, int, int]],
        content_pdfs: List[str],
        current_page: int
    ) -> int:
        """
        メインディレクトリ（大見出し）を処理

        Args:
            dir_path: ディレクトリのパス
            dir_name: ディレクトリ名
            create_separator_for_subfolder: サブフォルダに区切りページを作成するか
            toc_entries: 目次エントリのリスト（破壊的に更新される）
            content_pdfs: PDFパスのリスト（破壊的に更新される）
            current_page: 現在のページ番号

        Returns:
            int: 更新後のページ番号
        """
        heading = dir_name.strip()
        heading_for_toc = heading.replace("_", "")
        logger.info(f"=== メインディレクトリを処理中: {heading} ===")

        # 大見出し用の区切りページを作成
        logger.info(f"大見出し区切りページを作成: {heading}")
        sep_pdf = self.converter.create_separator_page(heading)
        if sep_pdf:
            content_pdfs.append(sep_pdf)
            toc_entries.append((heading_for_toc, PDFConstants.HEADING_LEVEL_MAIN, current_page))
            current_page += 1

        # サブディレクトリの処理
        subitems = sorted(os.listdir(dir_path))
        logger.info(f"ディレクトリ内のアイテム数: {len(subitems)}")
        for subitem in subitems:
            subitem_path = os.path.join(dir_path, subitem)

            if os.path.isdir(subitem_path):
                current_page = self._process_subfolder(
                    subitem_path, subitem, create_separator_for_subfolder,
                    toc_entries, content_pdfs, current_page
                )
            else:
                # ディレクトリ直下のファイル
                current_page = self._process_root_file(
                    subitem_path, subitem, toc_entries, content_pdfs, current_page
                )

        return current_page

    def _process_root_file(
        self,
        file_path: str,
        file_name: str,
        toc_entries: List[Tuple[str, int, int]],
        content_pdfs: List[str],
        current_page: int
    ) -> int:
        """
        ルートディレクトリ直下のファイルを処理

        Args:
            file_path: ファイルのパス
            file_name: ファイル名
            toc_entries: 目次エントリのリスト（破壊的に更新される）
            content_pdfs: PDFパスのリスト（破壊的に更新される）
            current_page: 現在のページ番号

        Returns:
            int: 更新後のページ番号
        """
        name = self._sanitize_name(os.path.splitext(file_name)[0])
        converted_pdf = self.converter.convert(file_path)
        if converted_pdf:
            content_pdfs.append(converted_pdf)
            toc_entries.append((name, PDFConstants.HEADING_LEVEL_SUB, current_page))
            current_page += self.processor.get_page_count(converted_pdf)
        return current_page

    def collect_documents(
        self,
        target_dir: str,
        create_separator_for_subfolder: bool = True
    ) -> Tuple[List[Tuple[str, int, int]], List[str]]:
        """
        ディレクトリを再帰的に探索し、ドキュメントを収集

        Args:
            target_dir: 探索対象のディレクトリ
            create_separator_for_subfolder: サブフォルダに区切りページを作成するか

        Returns:
            tuple: (目次エントリのリスト, 変換済みPDFパスのリスト)
        """
        toc_entries: List[Tuple[str, int, int]] = []
        content_pdfs: List[str] = []
        current_page = PDFConstants.CONTENT_START_PAGE

        items = sorted(os.listdir(target_dir))
        total_items = len(items)
        logger.info(f"ドキュメント収集を開始: {target_dir}")
        logger.info(f"処理対象アイテム数: {total_items}")

        for idx, item in enumerate(items, 1):
            if self.is_cancelled():
                logger.info("ドキュメント収集がキャンセルされました")
                raise CancelledError("ドキュメント収集がキャンセルされました")

            item_path = os.path.join(target_dir, item)
            logger.info(f"--- 処理中 [{idx}/{total_items}]: {item} ---")

            # 表紙ファイルの処理
            if os.path.isfile(item_path) and PDFConstants.COVER_FILE_KEYWORD in item:
                current_page = self._process_cover_file(item_path, content_pdfs, current_page)
                continue

            # ディレクトリの処理
            if os.path.isdir(item_path):
                current_page = self._process_directory(
                    item_path, item, create_separator_for_subfolder,
                    toc_entries, content_pdfs, current_page
                )
            else:
                # ルートディレクトリ直下のファイル（表紙以外）
                current_page = self._process_root_file(
                    item_path, item, toc_entries, content_pdfs, current_page
                )

        logger.info(f"ドキュメント収集完了: {len(content_pdfs)}ファイル, {len(toc_entries)}目次エントリ")

        # 空ディレクトリチェック
        if not content_pdfs:
            from exceptions import PDFProcessingError
            raise PDFProcessingError(
                f"処理可能なドキュメントが見つかりませんでした。\n\n"
                f"ディレクトリ: {target_dir}\n"
                f"サポートされているファイル形式:\n"
                f"  - Office: .doc, .docx, .xls, .xlsx, .ppt, .pptx\n"
                f"  - PDF: .pdf\n"
                f"  - 画像: .jpg, .jpeg, .png\n"
                f"  - 一太郎: .jtd",
                operation="ドキュメント収集"
            )

        return toc_entries, content_pdfs
