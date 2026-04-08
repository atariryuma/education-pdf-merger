"""
ドキュメント収集モジュール

ディレクトリ構造を探索し、ファイルをPDF化して目次を作成
"""
import logging
import os
from typing import List, Tuple, Callable, Optional

from core.pdf_converter import PDFConverter
from core.pdf_processor import PDFProcessor
from shared.exceptions import CancelledError, PDFProcessingError
from shared.constants import PDFConstants

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
        cancel_check: Optional[Callable[[], bool]] = None,
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
        self._collected_pdfs: List[str] = []
        self._cover_pages: int = PDFConstants.COVER_PAGE_COUNT

    def get_collected_pdfs(self) -> List[str]:
        """収集済みPDFのリストを返す（部分的な失敗回復用）"""
        return list(self._collected_pdfs)

    def get_cover_pages(self) -> int:
        """表紙の実際のページ数を返す"""
        return self._cover_pages

    def is_cancelled(self) -> bool:
        """キャンセルされたかどうかを確認"""
        return self._cancel_check()

    @staticmethod
    def _sanitize_name(name: str) -> str:
        """名前から先頭の数字とスペース、アンダースコアを除去"""
        return name.lstrip("0123456789 ").strip().replace("_", "")

    def _convert_and_add_pdf(
        self, file_path: str, content_pdfs: List[str], current_page: int
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

    def _add_separator(
        self,
        heading: str,
        level: int,
        toc_entries: List[Tuple[str, int, int]],
        content_pdfs: List[str],
        current_page: int,
    ) -> int:
        """
        区切りページを生成してコンテンツに追加し、目次エントリを登録する

        Args:
            heading: 区切りページの見出し
            level: 目次レベル（HEADING_LEVEL_MAIN / HEADING_LEVEL_SUB）
            toc_entries: 目次エントリのリスト（破壊的に更新される）
            content_pdfs: PDFパスのリスト（破壊的に更新される）
            current_page: 現在のページ番号

        Returns:
            int: 更新後のページ番号
        """
        sep_pdf = self.converter.create_separator_page(heading)
        if sep_pdf:
            sep_pages = self.processor.get_page_count(sep_pdf)
            content_pdfs.append(sep_pdf)
            toc_entries.append((heading, level, current_page))
            current_page += sep_pages
        else:
            logger.warning(f"区切りページの生成に失敗しました: {heading}")
            toc_entries.append((heading, level, current_page))
        return current_page

    def _process_cover_file(
        self, file_path: str, content_pdfs: List[str], current_page: int
    ) -> int:
        """
        表紙ファイルを処理

        CONTENT_START_PAGEは「表紙1ページ＋目次1ページ」前提の固定値なので、
        表紙が複数ページの場合は超過分だけ current_page を補正する。

        Args:
            file_path: 表紙ファイルのパス
            content_pdfs: PDFパスのリスト（破壊的に更新される）
            current_page: 現在のページ番号

        Returns:
            int: 更新後のページ番号
        """
        file_name = os.path.basename(file_path)
        logger.info(f"表紙ファイルを処理中: {file_name}")
        converted_pdf = self.converter.convert(file_path)
        if converted_pdf:
            cover_pages = self.processor.get_page_count(converted_pdf)
            content_pdfs.append(converted_pdf)
            self._cover_pages = cover_pages
            # CONTENT_START_PAGEは表紙1ページを前提としているため、
            # 超過分のみcurrent_pageに加算する
            extra = cover_pages - PDFConstants.COVER_PAGE_COUNT
            if extra > 0:
                current_page += extra
                logger.info(
                    "表紙が%dページ: ページオフセット+%d",
                    cover_pages, extra
                )
            logger.info(f"表紙変換成功: {file_name} ({cover_pages}ページ)")
        else:
            logger.warning(f"表紙変換スキップ: {file_name}")
        return current_page

    def _process_subfolder(
        self,
        subfolder_path: str,
        subfolder_name: str,
        create_separator: bool,
        toc_entries: List[Tuple[str, int, int]],
        content_pdfs: List[str],
        current_page: int,
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

        # サブフォルダ内のファイルを確認（ディレクトリ・一時ファイルは除外）
        all_items = sorted(os.listdir(subfolder_path))
        files = [
            f
            for f in all_items
            if os.path.isfile(os.path.join(subfolder_path, f))
            and not PDFConverter._is_temporary_file(f)
        ]
        total_files = len(files)
        logger.info(
            f"サブフォルダ内のファイル数: {total_files} (全アイテム: {len(all_items)})"
        )

        # 空フォルダの場合は目次エントリもスキップ
        if total_files == 0:
            logger.info(f"空フォルダのためスキップ: {subfolder_name}")
            return current_page

        if create_separator:
            logger.info(f"区切りページを作成: {sub_heading}")
            current_page = self._add_separator(
                sub_heading, PDFConstants.HEADING_LEVEL_SUB,
                toc_entries, content_pdfs, current_page
            )
        else:
            # 区切りページなしで目次にのみ登録
            toc_entries.append(
                (sub_heading, PDFConstants.HEADING_LEVEL_SUB, current_page)
            )

        # ファイルを処理
        for file_idx, filename in enumerate(files, 1):
            if self.is_cancelled():
                raise CancelledError("ドキュメント収集がキャンセルされました")
            logger.debug(f"  [{file_idx}/{total_files}] {filename}")
            file_path = os.path.join(subfolder_path, filename)
            current_page = self._convert_and_add_pdf(
                file_path, content_pdfs, current_page
            )

        return current_page

    def _process_directory(
        self,
        dir_path: str,
        dir_name: str,
        create_separator_for_subfolder: bool,
        toc_entries: List[Tuple[str, int, int]],
        content_pdfs: List[str],
        current_page: int,
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
        heading = self._sanitize_name(dir_name)
        logger.info(f"=== メインディレクトリを処理中: {heading} ===")

        # 大見出し用の区切りページを作成
        logger.info(f"大見出し区切りページを作成: {heading}")
        current_page = self._add_separator(
            heading, PDFConstants.HEADING_LEVEL_MAIN,
            toc_entries, content_pdfs, current_page
        )

        # サブディレクトリの処理
        subitems = sorted(os.listdir(dir_path))
        logger.info(f"ディレクトリ内のアイテム数: {len(subitems)}")
        for subitem in subitems:
            if self.is_cancelled():
                raise CancelledError("ドキュメント収集がキャンセルされました")
            subitem_path = os.path.join(dir_path, subitem)

            if os.path.isdir(subitem_path):
                current_page = self._process_subfolder(
                    subitem_path,
                    subitem,
                    create_separator_for_subfolder,
                    toc_entries,
                    content_pdfs,
                    current_page,
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
        current_page: int,
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
        self, target_dir: str, create_separator_for_subfolder: bool = True
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
        self._collected_pdfs = []
        self._cover_pages = PDFConstants.COVER_PAGE_COUNT
        content_pdfs = (
            self._collected_pdfs
        )  # 同一リスト参照: 例外時に部分結果を回収可能にする
        current_page = PDFConstants.CONTENT_START_PAGE

        all_items = sorted(os.listdir(target_dir))
        # 一時ファイル（~$, .$td, .$$$等）を収集段階で除外
        items = [
            item for item in all_items if not PDFConverter._is_temporary_file(item)
        ]
        if len(items) < len(all_items):
            logger.debug(f"一時ファイルを除外: {len(all_items) - len(items)}件")
        total_items = len(items)
        logger.info(f"ドキュメント収集を開始: {target_dir}")
        logger.info(f"処理対象アイテム数: {total_items}")

        for idx, item in enumerate(items, 1):
            if self.is_cancelled():
                logger.info("ドキュメント収集がキャンセルされました")
                raise CancelledError("ドキュメント収集がキャンセルされました")

            item_path = os.path.join(target_dir, item)
            logger.info(f"--- 処理中 [{idx}/{total_items}]: {item} ---")

            # 表紙ファイルの処理（ファイル名の先頭にキーワードが含まれるもの）
            if os.path.isfile(item_path) and item.startswith(
                PDFConstants.COVER_FILE_KEYWORD
            ):
                current_page = self._process_cover_file(
                    item_path, content_pdfs, current_page
                )
                continue

            # ディレクトリの処理
            if os.path.isdir(item_path):
                current_page = self._process_directory(
                    item_path,
                    item,
                    create_separator_for_subfolder,
                    toc_entries,
                    content_pdfs,
                    current_page,
                )
            else:
                # ルートディレクトリ直下のファイル（表紙以外）
                current_page = self._process_root_file(
                    item_path, item, toc_entries, content_pdfs, current_page
                )

        logger.info(
            f"ドキュメント収集完了: {len(content_pdfs)}ファイル, {len(toc_entries)}目次エントリ"
        )

        # 空ディレクトリチェック
        if not content_pdfs:
            raise PDFProcessingError(
                f"処理可能なドキュメントが見つかりませんでした。\n\n"
                f"ディレクトリ: {target_dir}\n"
                f"サポートされているファイル形式:\n"
                f"  - Office: .doc, .docx, .xls, .xlsx, .ppt, .pptx\n"
                f"  - PDF: .pdf\n"
                f"  - 画像: .jpg, .jpeg, .png\n"
                f"  - 一太郎: .jtd",
                operation="ドキュメント収集",
            )

        return toc_entries, content_pdfs
