"""
Excel転記処理の共通基底クラス

GoogleSheetsTransferとExcelTransferの共通機能を提供
"""
import logging
from abc import ABC, abstractmethod
from datetime import datetime
from typing import Any, Dict, Tuple, Optional, Callable

try:
    import win32com.client
    import pythoncom
except ImportError as e:
    raise ImportError(
        "pywin32 がインストールされていません。\n"
        "以下のコマンドを実行してインストールしてください:\n"
        "pip install pywin32"
    ) from e

from constants import ExcelSortOrder, ExcelSortHeader, ExcelTransferConstants, PDFConversionConstants

# ロガーの設定
logger = logging.getLogger(__name__)


class BaseExcelTransfer(ABC):
    """Excel転記処理の共通基底クラス"""

    def __init__(
        self,
        target_filename: str,
        target_sheet: str,
        progress_callback: Optional[Callable[[str], None]] = None,
        cancel_check: Optional[Callable[[], bool]] = None
    ) -> None:
        """
        初期化

        Args:
            target_filename: ターゲットExcelファイル名
            target_sheet: ターゲットシート名
            progress_callback: 進捗状況を報告するコールバック関数
            cancel_check: キャンセル確認用コールバック関数（Trueで中断）
        """
        self.target_filename = target_filename
        self.target_sheet = target_sheet
        self.progress_callback = progress_callback
        self.cancel_check = cancel_check

        # Excel COMオブジェクト
        self.excel: Any = None
        self.target_wb: Any = None
        self.target_ws: Any = None
        self._com_initialized: bool = False

    def _report_progress(self, message: str) -> None:
        """
        進捗状況を報告

        Args:
            message: 進捗メッセージ
        """
        logger.info(message)
        if self.progress_callback:
            self.progress_callback(message)

    def _check_cancelled(self) -> bool:
        """
        キャンセルチェック

        Returns:
            bool: キャンセルされた場合True
        """
        if self.cancel_check and self.cancel_check():
            logger.info("ユーザーによって処理がキャンセルされました")
            return True
        return False

    @abstractmethod
    def _read_data_row(self, found_row: int, start_col: str, end_col: str) -> list:
        """
        データソースから行データを読み取る（抽象メソッド）

        Args:
            found_row: 行番号
            start_col: 開始列
            end_col: 終了列

        Returns:
            list: 行データ
        """
        pass

    @abstractmethod
    def _find_value_in_source(self, search_value: str) -> Optional[int]:
        """
        データソースから値を検索（抽象メソッド）

        Args:
            search_value: 検索値

        Returns:
            Optional[int]: 見つかった行番号（見つからない場合None）
        """
        pass

    @abstractmethod
    def _read_cell_value(self, row: int, col: str) -> Any:
        """
        データソースからセル値を読み取る（抽象メソッド）

        Args:
            row: 行番号
            col: 列

        Returns:
            Any: セル値
        """
        pass

    @abstractmethod
    def _get_error_class(self) -> type:
        """
        エラークラスを取得（抽象メソッド）

        Returns:
            type: 使用するエラークラス
        """
        pass

    def _count_events_in_found_row(
        self,
        found_row: int,
        filter_keyword: Optional[str] = None
    ) -> Dict[int, Tuple[int, int]]:
        """
        データソースの found_row の【E～AN】列を一括で取得し、学年毎にカウント

        Args:
            found_row: データソースの行番号
            filter_keyword: フィルター用キーワード（D8～D50用、Noneの場合はループ2,3）

        Returns:
            辞書 {grade: (event_count, absent_count), ...}（grade 1～6）

        Raises:
            SystemExit, KeyboardInterrupt: 即座に再スロー
        """
        error_class = self._get_error_class()
        try:
            # E～AN列を一括取得
            row_data = self._read_data_row(
                found_row,
                ExcelTransferConstants.REF_DATA_START_COL,
                ExcelTransferConstants.REF_DATA_END_COL
            )

            if not row_data:
                logger.warning(f"行 {found_row} のデータが空です")
                return {
                    grade: (0, 0)
                    for grade in range(1, ExcelTransferConstants.GRADES_COUNT + 1)
                }

            # データ長の検証
            if len(row_data) < ExcelTransferConstants.TOTAL_COLUMNS:
                logger.warning(
                    f"行 {found_row} のデータが不足しています"
                    f"（期待: {ExcelTransferConstants.TOTAL_COLUMNS}列、実際: {len(row_data)}列）"
                )

            counts = {}

            for grade in range(1, ExcelTransferConstants.GRADES_COUNT + 1):
                start_index = (grade - 1) * ExcelTransferConstants.PERIODS_PER_GRADE
                end_index = start_index + ExcelTransferConstants.PERIODS_PER_GRADE

                # 範囲外チェック
                if end_index > len(row_data):
                    counts[grade] = (0, 0)
                    continue

                group = row_data[start_index:end_index]

                event_count = 0
                absent_count = 0

                if filter_keyword is not None:
                    # D8～D50用：完全一致のみカウント
                    if filter_keyword == ExcelTransferConstants.ABSENT_KEYWORD:
                        absent_count = sum(
                            1 for cell in group
                            if cell is not None and str(cell).strip() == ExcelTransferConstants.ABSENT_KEYWORD
                        )
                        event_count = 0
                    else:
                        event_count = sum(
                            1 for cell in group
                            if cell is not None and str(cell).strip() == filter_keyword
                        )
                        absent_count = 0
                else:
                    # ループ2,3用：部分一致でカウント
                    event_count = sum(
                        1 for cell in group
                        if cell is not None and any(
                            keyword in str(cell)
                            for keyword in ExcelTransferConstants.EVENT_KEYWORDS
                        )
                    )
                    absent_count = sum(
                        1 for cell in group
                        if cell is not None and ExcelTransferConstants.ABSENT_KEYWORD in str(cell)
                    )

                counts[grade] = (event_count, absent_count)

            return counts

        except (SystemExit, KeyboardInterrupt):
            # システム終了・キーボード割り込みは即座に再スロー
            raise
        except Exception as e:
            logger.error(f"行事時数・欠時数のカウント中にエラー（行 {found_row}）: {e}")
            raise error_class(
                f"行 {found_row} のデータ処理中にエラーが発生しました。\n"
                f"データに問題がある可能性があります。\n"
                f"詳細: {e}",
                operation="データカウント",
                original_error=e
            ) from e

    def _process_row(
        self,
        row: int,
        search_col: str,
        filter_keyword: Optional[str] = None
    ) -> None:
        """
        1行分の転記処理（データソース → Excel）

        Args:
            row: ターゲットExcelファイルの行番号
            search_col: 検索値を取得する列（D or C）
            filter_keyword: フィルター用キーワード（Noneの場合は部分一致）

        Note:
            このメソッドはターゲットExcelワークシートを直接変更します（副作用あり）
        """
        error_class = self._get_error_class()

        # キャンセルチェック
        if self._check_cancelled():
            raise error_class("ユーザーによって処理がキャンセルされました", operation="転記処理")

        # 検索値を取得
        search_value = self.target_ws.Range(f"{search_col}{row}").Value
        if search_value is None:
            return

        # データソースから検索
        found_row = self._find_value_in_source(search_value)

        if found_row is not None:
            # 日付を転記
            ref_date = self._read_cell_value(
                found_row,
                ExcelTransferConstants.TARGET_DATE_COL
            )
            self.target_ws.Range(
                f"{ExcelTransferConstants.TARGET_DATE_COL}{row}"
            ).Value = ref_date

            # 処理日を記録
            today_str = datetime.today().strftime('%Y-%m-%d')
            self.target_ws.Range(
                f"{ExcelTransferConstants.TARGET_PROCESS_DATE_COL}{row}"
            ).Value = today_str

            logger.debug(
                f"Row {row} ({search_col}列): '{search_value}' → "
                f"データソース行 {found_row} の A列: {ref_date}"
            )

            # 行事時数・欠時数をカウント
            counts = self._count_events_in_found_row(found_row, filter_keyword)

            # 学年別に転記
            for grade in range(1, ExcelTransferConstants.GRADES_COUNT + 1):
                event_count, absent_count = counts[grade]
                tgt_event, tgt_absent = ExcelTransferConstants.GRADE_COLUMN_MAPPING[grade]

                # 0の場合は空白にする（帳票向け）
                self.target_ws.Range(f"{tgt_event}{row}").Value = event_count if event_count else ""
                self.target_ws.Range(f"{tgt_absent}{row}").Value = absent_count if absent_count else ""

                logger.debug(
                    f"  → Grade {grade} (Row {row}): "
                    f"行事時数={event_count}, 欠時数={absent_count}"
                )
        else:
            # 見つからない場合はA列をクリア
            self.target_ws.Range(
                f"{ExcelTransferConstants.TARGET_DATE_COL}{row}"
            ).Value = ""
            logger.debug(
                f"Row {row} ({search_col}列): '{search_value}' はデータソースに見つかりませんでした"
            )

    def _sort_range(self, range_str: str, key_cell: str) -> None:
        """
        指定範囲をB列（日付）で昇順ソート

        Args:
            range_str: ソート範囲（例: "A8:P50"）
            key_cell: ソートキー（例: "B8"）

        Raises:
            エラークラス: 並び替えに失敗した場合
        """
        error_class = self._get_error_class()
        try:
            # 結合セルがある場合は解除
            try:
                self.target_ws.Range(range_str).UnMerge()
            except Exception as e:
                # 結合セルがない場合のエラーは無視
                logger.debug(f"UnMerge実行（結合セルなしの可能性）: {e}")

            # B列の日付順にソート（昇順）
            self.target_ws.Range(range_str).Sort(
                Key1=self.target_ws.Range(key_cell),
                Order1=ExcelSortOrder.ASCENDING,
                Header=ExcelSortHeader.NO
            )
            logger.debug(f"範囲 {range_str} を日付順に並び替えました")
        except Exception as e:
            logger.error(f"並び替え中にエラー ({range_str}): {e}")
            raise error_class(
                f"データの並び替えに失敗しました。\n"
                f"範囲: {range_str}\n"
                f"詳細: {e}\n\n"
                f"結合セルまたはデータ形式に問題がある可能性があります。",
                operation="並び替え",
                original_error=e
            ) from e

    def _execute_transfer_loops(self) -> None:
        """
        3つの転記ループを実行（共通処理）

        Raises:
            エラークラス: 転記処理中にエラーが発生した場合
        """
        error_class = self._get_error_class()

        # ループ1: D8～D50（フィルターあり）
        logger.info(PDFConversionConstants.LOG_SEPARATOR_MINOR)
        logger.info(
            f"【ループ1】{ExcelTransferConstants.LOOP1_SEARCH_COL}"
            f"{ExcelTransferConstants.LOOP1_START_ROW}～"
            f"{ExcelTransferConstants.LOOP1_END_ROW - 1} の処理を開始"
        )
        self._report_progress("ループ1: フィルター付き転記を実行中...")

        # C列のフィルターキーワードを一括取得
        start_row = ExcelTransferConstants.LOOP1_START_ROW
        end_row = ExcelTransferConstants.LOOP1_END_ROW
        filter_range_addr = (
            f"{ExcelTransferConstants.TARGET_FILTER_COL}{start_row}:"
            f"{ExcelTransferConstants.TARGET_FILTER_COL}{end_row - 1}"
        )
        filter_range_values = self.target_ws.Range(filter_range_addr).Value

        # 一括取得した値をリスト化
        filter_list = []
        if filter_range_values:
            if isinstance(filter_range_values, tuple):
                filter_list = [row[0] if row else None for row in filter_range_values]
            else:
                filter_list = [filter_range_values]

        logger.debug(f"フィルターキーワードを一括取得: {len(filter_list)}件")

        for i, row in enumerate(range(start_row, end_row)):
            filter_keyword = filter_list[i] if i < len(filter_list) else None
            if filter_keyword is not None:
                filter_keyword = str(filter_keyword).strip()

            self._process_row(row, ExcelTransferConstants.LOOP1_SEARCH_COL, filter_keyword)

        # ループ1の範囲を並び替え
        self._sort_range(
            ExcelTransferConstants.LOOP1_SORT_RANGE,
            ExcelTransferConstants.LOOP1_SORT_KEY
        )
        logger.info("【ループ1】完了")

        # ループ2: C55～C62（フィルターなし）
        logger.info(PDFConversionConstants.LOG_SEPARATOR_MINOR)
        logger.info(
            f"【ループ2】{ExcelTransferConstants.LOOP2_SEARCH_COL}"
            f"{ExcelTransferConstants.LOOP2_START_ROW}～"
            f"{ExcelTransferConstants.LOOP2_END_ROW - 1} の処理を開始"
        )
        self._report_progress("ループ2: 通常転記を実行中...")

        for row in range(ExcelTransferConstants.LOOP2_START_ROW,
                       ExcelTransferConstants.LOOP2_END_ROW):
            self._process_row(row, ExcelTransferConstants.LOOP2_SEARCH_COL, None)

        # ループ2の範囲を並び替え
        self._sort_range(
            ExcelTransferConstants.LOOP2_SORT_RANGE,
            ExcelTransferConstants.LOOP2_SORT_KEY
        )
        logger.info("【ループ2】完了")

        # ループ3: C67～C96（フィルターなし）
        logger.info(PDFConversionConstants.LOG_SEPARATOR_MINOR)
        logger.info(
            f"【ループ3】{ExcelTransferConstants.LOOP3_SEARCH_COL}"
            f"{ExcelTransferConstants.LOOP3_START_ROW}～"
            f"{ExcelTransferConstants.LOOP3_END_ROW - 1} の処理を開始"
        )
        self._report_progress("ループ3: 通常転記を実行中...")

        for row in range(ExcelTransferConstants.LOOP3_START_ROW,
                       ExcelTransferConstants.LOOP3_END_ROW):
            self._process_row(row, ExcelTransferConstants.LOOP3_SEARCH_COL, None)

        # ループ3の範囲を並び替え
        self._sort_range(
            ExcelTransferConstants.LOOP3_SORT_RANGE,
            ExcelTransferConstants.LOOP3_SORT_KEY
        )
        logger.info("【ループ3】完了")

    @abstractmethod
    def execute(self) -> None:
        """
        転記処理を実行（抽象メソッド）

        サブクラスで実装し、接続・転記・クリーンアップを行う
        """
        pass
