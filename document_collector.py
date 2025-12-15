"""
ドキュメント収集モジュール

ディレクトリ構造を探索し、ファイルをPDF化して目次を作成
"""
import logging
import os
from typing import List, Tuple, Callable, Optional

from config_loader import ConfigLoader
from pdf_converter import PDFConverter
from pdf_processor import PDFProcessor
from exceptions import PDFProcessingError, CancelledError

# ロガーの設定
logger = logging.getLogger(__name__)

# マジックナンバーの定数化
COVER_PAGE_COUNT = 1  # 表紙のページ数
TOC_PAGE_COUNT = 1    # 目次のページ数
INITIAL_CONTENT_PAGE = COVER_PAGE_COUNT + TOC_PAGE_COUNT + 1  # 最初のコンテンツページ番号


class DocumentCollector:
    """ディレクトリを探索してドキュメントを収集し、目次を生成するクラス"""

    # 目次見出しレベル
    HEADING_LEVEL_MAIN = 1
    HEADING_LEVEL_SUB = 2

    def __init__(
        self,
        pdf_converter: PDFConverter,
        pdf_processor: PDFProcessor,
        template_path: str,
        cancel_check: Optional[Callable[[], bool]] = None
    ) -> None:
        """
        Args:
            pdf_converter: PDFConverterインスタンス
            pdf_processor: PDFProcessorインスタンス
            template_path: 区切りページテンプレートのパス
            cancel_check: キャンセル状態をチェックするコールバック関数
        """
        self.converter = pdf_converter
        self.processor = pdf_processor
        self.template_path = template_path
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
        """ファイルをPDFに変換してリストに追加し、更新されたページ番号を返す"""
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
        """表紙ファイルを処理"""
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
        """サブフォルダを処理"""
        sub_heading = self._sanitize_name(subfolder_name)
        logger.info(f"サブフォルダを処理中: {subfolder_name}")

        if create_separator:
            # サブフォルダにも区切りページを作成
            logger.info(f"区切りページを作成: {sub_heading}")
            sep_pdf = self.converter.create_separator_page(sub_heading, self.template_path)
            if sep_pdf:
                content_pdfs.append(sep_pdf)
                toc_entries.append((sub_heading, self.HEADING_LEVEL_SUB, current_page))
                current_page += 1
        else:
            # 区切りページなしで目次にのみ登録
            toc_entries.append((sub_heading, self.HEADING_LEVEL_SUB, current_page))

        # サブフォルダ内のファイルを処理
        files = sorted(os.listdir(subfolder_path))
        logger.info(f"サブフォルダ内のファイル数: {len(files)}")
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
        """メインディレクトリ（大見出し）を処理"""
        heading = dir_name.strip()
        heading_for_toc = heading.replace("_", "")
        logger.info(f"=== メインディレクトリを処理中: {heading} ===")

        # 大見出し用の区切りページを作成
        logger.info(f"大見出し区切りページを作成: {heading}")
        sep_pdf = self.converter.create_separator_page(heading, self.template_path)
        if sep_pdf:
            content_pdfs.append(sep_pdf)
            toc_entries.append((heading_for_toc, self.HEADING_LEVEL_MAIN, current_page))
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
                name = self._sanitize_name(os.path.splitext(subitem)[0])
                logger.info(f"ファイルを変換中: {subitem}")
                converted_pdf = self.converter.convert(subitem_path)
                if converted_pdf:
                    page_count = self.processor.get_page_count(converted_pdf)
                    content_pdfs.append(converted_pdf)
                    toc_entries.append((name, self.HEADING_LEVEL_SUB, current_page))
                    logger.info(f"変換成功: {subitem} ({page_count}ページ)")
                    current_page += page_count

        return current_page

    def _process_root_file(
        self,
        file_path: str,
        file_name: str,
        toc_entries: List[Tuple[str, int, int]],
        content_pdfs: List[str],
        current_page: int
    ) -> int:
        """ルートディレクトリ直下のファイルを処理"""
        name = self._sanitize_name(os.path.splitext(file_name)[0])
        converted_pdf = self.converter.convert(file_path)
        if converted_pdf:
            content_pdfs.append(converted_pdf)
            toc_entries.append((name, self.HEADING_LEVEL_SUB, current_page))
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
        current_page = INITIAL_CONTENT_PAGE

        items = sorted(os.listdir(target_dir))
        total_items = len(items)
        logger.info(f"ドキュメント収集を開始: {target_dir}")
        logger.info(f"処理対象アイテム数: {total_items}")

        for idx, item in enumerate(items, 1):
            if self.is_cancelled():
                logger.info("ドキュメント収集がキャンセルされました")
                return toc_entries, content_pdfs

            item_path = os.path.join(target_dir, item)
            logger.info(f"--- 処理中 [{idx}/{total_items}]: {item} ---")

            # 表紙ファイルの処理
            if os.path.isfile(item_path) and "表紙" in item:
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
        return toc_entries, content_pdfs


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
        """
        logger.info(f"PDFマージ処理を開始: {target_dir}")
        logger.info(f"出力先: {output_pdf}")

        # 1. ドキュメント収集とPDF変換
        logger.info("[Step 1/6] ドキュメントを収集・変換中...")
        toc_entries, content_pdfs = self.collector.collect_documents(
            target_dir,
            create_separator_for_subfolder
        )
        self._check_cancel()

        # 2. 一時的にマージ
        logger.info("[Step 2/6] 一時マージPDFを作成中...")
        temp_merged = os.path.join(self.temp_dir, "temp_merged.pdf")
        self.processor.merge_pdfs(content_pdfs, temp_merged)
        self._check_cancel()

        # 3. 目次PDFを生成
        logger.info("[Step 3/6] 目次を作成中...")
        toc_pdf = os.path.join(self.temp_dir, "toc.pdf")
        self.processor.create_toc_pdf(toc_entries, toc_pdf)
        self._check_cancel()

        # 4. 表紙と残りのページに分割
        logger.info("[Step 4/6] 表紙とコンテンツを分割中...")
        cover_pdf, remainder_pdf = self.processor.split_pdf(temp_merged, self.temp_dir)
        self._check_cancel()

        # 分割結果の検証
        if not cover_pdf or not os.path.exists(cover_pdf):
            raise PDFProcessingError("分割", "表紙PDFが作成されませんでした")
        if not remainder_pdf or not os.path.exists(remainder_pdf):
            raise PDFProcessingError("分割", "コンテンツPDFが作成されませんでした")

        # 5. 最終的にマージ（表紙 + 目次 + 残り）
        logger.info("[Step 5/6] 最終PDFをマージ中...")
        final_list = [cover_pdf, toc_pdf, remainder_pdf]
        self.processor.merge_pdfs(final_list, output_pdf)
        self._check_cancel()

        # 6. ページ番号を追加（先頭ページは除外）
        logger.info("[Step 6/6] ページ番号としおりを追加中...")
        self.processor.add_page_numbers(output_pdf, exclude_first_pages=COVER_PAGE_COUNT)
        self._check_cancel()

        # 7. PDFアウトライン（しおり）を設定
        self.processor.set_pdf_outlines(output_pdf, toc_entries)

        total_pages = self.processor.get_page_count(output_pdf)
        logger.info(f"PDFの作成が完了しました: {output_pdf}")
        logger.info(f"  目次エントリ数: {len(toc_entries)}")
        logger.info(f"  総ページ数: {total_pages}")
