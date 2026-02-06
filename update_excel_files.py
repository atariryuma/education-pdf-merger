"""
Excel自動転記モジュール

年間行事計画（参照ファイル）から様式ファイル（反映ファイル）へ
日付、行事時数、欠時数を自動転記・集計する
"""
import logging
import os
from datetime import datetime
from typing import Any, Dict, Tuple, Optional, Callable, List

try:
    import win32com.client
    import pythoncom
except ImportError as e:
    raise ImportError(
        "pywin32 がインストールされていません。\n"
        "以下のコマンドを実行してインストールしてください:\n"
        "pip install pywin32"
    ) from e

from exceptions import PDFMergeError
from constants import ExcelLookIn, ExcelLookAt, ExcelSortOrder, ExcelSortHeader, ExcelTransferConstants, PDFConversionConstants

# ロガーの設定
logger = logging.getLogger(__name__)


class ExcelTransferError(PDFMergeError):
    """Excel転記処理エラー"""
    pass


class ExcelTransfer:
    """Excel自動転記処理クラス"""

    def __init__(
        self,
        ref_filename: str,
        target_filename: str,
        ref_sheet: str,
        target_sheet: str,
        progress_callback: Optional[Callable[[str], None]] = None,
        cancel_check: Optional[Callable[[], bool]] = None
    ) -> None:
        """
        初期化

        Args:
            ref_filename: 参照ファイル名
            target_filename: 反映ファイル名
            ref_sheet: 参照シート名
            target_sheet: 反映シート名
            progress_callback: 進捗状況を報告するコールバック関数
            cancel_check: キャンセル確認用コールバック関数（Trueで中断）
        """
        self.ref_filename = ref_filename
        self.target_filename = target_filename
        self.ref_sheet = ref_sheet
        self.target_sheet = target_sheet
        self.progress_callback = progress_callback
        self.cancel_check = cancel_check

        # Excel COMオブジェクト
        self.excel: Any = None
        self.ref_wb: Any = None
        self.ref_ws: Any = None
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

    def _check_cancelled(self) -> None:
        """
        キャンセルチェック

        Raises:
            ExcelTransferError: キャンセルされた場合
        """
        if self.cancel_check and self.cancel_check():
            logger.info("ユーザーによって処理がキャンセルされました")
            raise ExcelTransferError("ユーザーによって処理がキャンセルされました", operation="転記処理")

    def _read_data_row(self, found_row: int, start_col: str, end_col: str) -> list:
        """
        参照Excelから行データを読み取る

        Args:
            found_row: 行番号
            start_col: 開始列
            end_col: 終了列

        Returns:
            list: 行データ（文字列に変換済み）
        """
        range_str = f"{start_col}{found_row}:{end_col}{found_row}"
        rng = self.ref_ws.Range(range_str).Value

        if rng is None:
            return []

        # Excelは2次元タプルで返すので、1次元リストに変換
        # 各セル値を安全に文字列化（日付・整数などの予期しない型に対応）
        if isinstance(rng, tuple):
            return [str(cell) if cell is not None else "" for cell in rng[0]]
        else:
            return [str(rng) if rng is not None else ""]

    def _find_value_in_source(self, search_value: str) -> Optional[int]:
        """
        参照ExcelのC列から値を検索

        Args:
            search_value: 検索値

        Returns:
            Optional[int]: 見つかった行番号（見つからない場合None）
        """
        # C列の使用範囲内で検索（Columns()は列全体で遅く、エラーになる可能性がある）
        ref_col_C = self.ref_ws.Range("C:C")
        found_cell = ref_col_C.Find(
            What=search_value,
            LookIn=ExcelLookIn.VALUES,
            LookAt=ExcelLookAt.PART
        )

        return found_cell.Row if found_cell is not None else None

    def _read_cell_value(self, row: int, col: str) -> Any:
        """
        参照Excelからセル値を読み取る

        Args:
            row: 行番号
            col: 列

        Returns:
            Any: セル値
        """
        # 列が文字列（"A"）の場合は列番号に変換
        if isinstance(col, str) and len(col) == 1:
            col_num = ord(col.upper()) - ord('A') + 1
            return self.ref_ws.Cells(row, col_num).Value
        else:
            # 列名指定の場合
            return self.ref_ws.Range(f"{col}{row}").Value

    def _count_events_in_found_row(
        self,
        found_row: int,
        filter_keyword: Optional[str] = None
    ) -> Dict[int, Tuple[int, int]]:
        """
        参照Excelの found_row の【E～AN】列を一括で取得し、学年毎にカウント

        Args:
            found_row: 参照Excelの行番号
            filter_keyword: フィルター用キーワード（D8～D50用、Noneの場合はループ2,3）

        Returns:
            辞書 {grade: (event_count, absent_count), ...}（grade 1～6）

        Raises:
            SystemExit, KeyboardInterrupt: 即座に再スロー
        """
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
            raise ExcelTransferError(
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
        1行分の転記処理（参照Excel → ターゲットExcel）

        Args:
            row: ターゲットExcelファイルの行番号
            search_col: 検索値を取得する列（D or C）
            filter_keyword: フィルター用キーワード（Noneの場合は部分一致）

        Note:
            このメソッドはターゲットExcelワークシートを直接変更します（副作用あり）
        """
        # キャンセルチェック
        self._check_cancelled()

        # 検索値を取得
        search_value = self.target_ws.Range(f"{search_col}{row}").Value
        if search_value is None:
            logger.info(f"  - 行{row}: {search_col}列が空白のためスキップ")
            return

        logger.info(f"  > 行{row}: {search_col}列の値 '{search_value}' を参照Excelで検索中...")

        # 参照Excelから検索
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

            # デバッグ情報をINFOレベルで出力
            logger.info(
                f"  ✓ 行{row}: '{search_value}' → 参照Excel行{found_row} (日付: {ref_date})"
            )

            logger.debug(
                f"Row {row} ({search_col}列): '{search_value}' → "
                f"参照Excel行 {found_row} の A列: {ref_date}"
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
            logger.info(f"  ⚠ 行{row}: '{search_value}' → 参照Excelに見つかりませんでした")
            logger.debug(
                f"Row {row} ({search_col}列): '{search_value}' は参照Excelに見つかりませんでした"
            )

    def _sort_range(self, range_str: str, key_cell: str) -> None:
        """
        指定範囲をB列（日付）で昇順ソート（可能な限りセル結合を保持）

        Args:
            range_str: ソート範囲（例: "A8:P50"）
            key_cell: ソートキー（例: "B8"）

        Raises:
            ExcelTransferError: 並び替えに失敗した場合
        """
        try:
            # まずUnMerge無しでソートを試行
            try:
                self.target_ws.Range(range_str).Sort(
                    Key1=self.target_ws.Range(key_cell),
                    Order1=ExcelSortOrder.ASCENDING,
                    Header=ExcelSortHeader.NO
                )
                logger.info(f"範囲 {range_str} をセル結合を保持したまま並び替えました")
            except Exception as sort_error:
                # ソート失敗時はUnMerge して再試行
                logger.warning(
                    f"セル結合があるためUnMergeしてソートします: {sort_error}"
                )
                try:
                    self.target_ws.Range(range_str).UnMerge()
                except Exception as unmerge_error:
                    # 結合セルがない場合のエラーは無視
                    logger.debug(f"UnMerge実行（結合セルなしの可能性）: {unmerge_error}")

                # 再度ソート実行
                self.target_ws.Range(range_str).Sort(
                    Key1=self.target_ws.Range(key_cell),
                    Order1=ExcelSortOrder.ASCENDING,
                    Header=ExcelSortHeader.NO
                )
                logger.info(f"範囲 {range_str} を並び替えました（セル結合解除）")

        except Exception as e:
            logger.error(f"並び替え中にエラー ({range_str}): {e}")
            raise ExcelTransferError(
                f"データの並び替えに失敗しました。\n"
                f"範囲: {range_str}\n"
                f"詳細: {e}\n\n"
                f"結合セルまたはデータ形式に問題がある可能性があります。",
                operation="並び替え",
                original_error=e
            ) from e

    def _execute_transfer_loops(self) -> None:
        """
        3つの転記ループを実行

        Raises:
            ExcelTransferError: 転記処理中にエラーが発生した場合
        """
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
        logger.info(f"ループ1: 行{start_row}～{end_row - 1}を処理します（全{end_row - start_row}行）")

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

        start_row = ExcelTransferConstants.LOOP2_START_ROW
        end_row = ExcelTransferConstants.LOOP2_END_ROW
        logger.info(f"ループ2: 行{start_row}～{end_row - 1}を処理します（全{end_row - start_row}行）")

        for row in range(start_row, end_row):
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

        start_row = ExcelTransferConstants.LOOP3_START_ROW
        end_row = ExcelTransferConstants.LOOP3_END_ROW
        logger.info(f"ループ3: 行{start_row}～{end_row - 1}を処理します（全{end_row - start_row}行）")

        for row in range(start_row, end_row):
            self._process_row(row, ExcelTransferConstants.LOOP3_SEARCH_COL, None)

        # ループ3の範囲を並び替え
        self._sort_range(
            ExcelTransferConstants.LOOP3_SORT_RANGE,
            ExcelTransferConstants.LOOP3_SORT_KEY
        )
        logger.info("【ループ3】完了")

    def _connect_to_excel(self) -> None:
        """
        既存のExcelインスタンスに接続してワークブック/シートを取得

        Raises:
            ExcelTransferError: 接続に失敗した場合
        """
        try:
            # COM初期化（スレッドごとに必要）
            try:
                pythoncom.CoInitialize()
                self._com_initialized = True
                logger.debug("COM初期化完了")
            except Exception as e:
                logger.debug(f"COM初期化スキップ（既に初期化済み）: {e}")
                self._com_initialized = False

            # 既存のExcelインスタンスを取得
            self.excel = win32com.client.Dispatch("Excel.Application")
            logger.debug("Excelインスタンスに接続しました")

            # ワークブックを取得（ファイル名で検索）
            self.ref_wb = None
            self.target_wb = None

            # フルパスの場合はファイル名のみを抽出
            ref_basename = os.path.basename(self.ref_filename)
            target_basename = os.path.basename(self.target_filename)

            for wb in self.excel.Workbooks:
                if ref_basename == wb.Name:
                    self.ref_wb = wb
                    logger.debug(f"参照ファイルを検出: {wb.Name}")
                if target_basename == wb.Name:
                    self.target_wb = wb
                    logger.debug(f"反映ファイルを検出: {wb.Name}")

            # ワークブックが見つからない場合のエラー
            if self.ref_wb is None:
                raise ExcelTransferError(
                    f"参照ファイルが開かれていません: {ref_basename}\n\n"
                    "Excelで該当ファイルを開いてから実行してください。"
                )
            if self.target_wb is None:
                raise ExcelTransferError(
                    f"反映ファイルが開かれていません: {target_basename}\n\n"
                    "Excelで該当ファイルを開いてから実行してください。"
                )

            # シートを取得
            try:
                # 参照ファイルの利用可能なシート名を取得
                ref_sheet_names = [ws.Name for ws in self.ref_wb.Worksheets]
                logger.info(f"参照ファイルの利用可能なシート: {ref_sheet_names}")
                logger.info(f"検索するシート名: '{self.ref_sheet}'")

                self.ref_ws = self.ref_wb.Worksheets(self.ref_sheet)
                logger.info(f"✓ 参照シートに接続: {self.ref_sheet}")
            except Exception as e:
                ref_sheet_names = [ws.Name for ws in self.ref_wb.Worksheets]
                raise ExcelTransferError(
                    f"参照シートが見つかりません: '{self.ref_sheet}'\n\n"
                    f"ファイル: {self.ref_filename}\n"
                    f"利用可能なシート: {ref_sheet_names}",
                    operation="Excel接続",
                    original_error=e
                ) from e

            try:
                # 対象ファイルの利用可能なシート名を取得
                target_sheet_names = [ws.Name for ws in self.target_wb.Worksheets]
                logger.info(f"対象ファイルの利用可能なシート: {target_sheet_names}")
                logger.info(f"検索するシート名: '{self.target_sheet}'")

                self.target_ws = self.target_wb.Worksheets(self.target_sheet)
                logger.info(f"✓ 対象シートに接続: {self.target_sheet}")
            except Exception as e:
                target_sheet_names = [ws.Name for ws in self.target_wb.Worksheets]
                raise ExcelTransferError(
                    f"対象シートが見つかりません: '{self.target_sheet}'\n\n"
                    f"ファイル: {self.target_filename}\n"
                    f"利用可能なシート: {target_sheet_names}",
                    operation="Excel接続",
                    original_error=e
                ) from e

            # 接続状態の検証
            if self.excel is None or self.ref_ws is None or self.target_ws is None:
                raise ExcelTransferError(
                    "Excelへの接続に失敗しました。\n"
                    "ワークシートまたはアプリケーションオブジェクトが取得できませんでした。",
                    operation="Excel接続"
                )

            # データ存在確認（サンプル値をログ出力）
            logger.info("--- データ存在確認 ---")
            sample_d8 = self.target_ws.Range("D8").Value
            sample_c55 = self.target_ws.Range("C55").Value
            sample_c67 = self.target_ws.Range("C67").Value
            logger.info(f"対象シートのサンプル値: D8='{sample_d8}', C55='{sample_c55}', C67='{sample_c67}'")

        except ExcelTransferError:
            raise
        except Exception as e:
            raise ExcelTransferError(
                f"Excelへの接続に失敗しました: {e}",
                operation="Excel接続",
                original_error=e
            ) from e

    def _save_target_workbook(self) -> None:
        """
        対象ワークブックを保存

        Note:
            変更内容をファイルに反映します
        """
        if self.target_wb is not None:
            try:
                logger.info("変更内容を保存中...")
                self.target_wb.Save()
                logger.info("✓ 変更内容を保存しました")
                self._report_progress("変更内容を保存しました")
            except Exception as e:
                logger.error(f"ファイル保存エラー: {e}", exc_info=True)
                raise ExcelTransferError(
                    f"Excelファイルの保存に失敗しました: {e}",
                    operation="ファイル保存",
                    original_error=e
                ) from e

    def _connect_to_target_only(self) -> None:
        """
        ターゲットファイルのみに接続（行事名読み込み専用）

        Raises:
            ExcelTransferError: 接続に失敗した場合
        """
        try:
            # COM初期化（スレッドごとに必要）
            try:
                pythoncom.CoInitialize()
                self._com_initialized = True
                logger.debug("COM初期化完了")
            except Exception as e:
                logger.debug(f"COM初期化スキップ（既に初期化済み）: {e}")
                self._com_initialized = False

            # 既存のExcelインスタンスを取得
            self.excel = win32com.client.Dispatch("Excel.Application")
            logger.debug("Excelインスタンスに接続しました")

            # ターゲットワークブックを取得
            self.target_wb = None
            target_basename = os.path.basename(self.target_filename)

            for wb in self.excel.Workbooks:
                if target_basename == wb.Name:
                    self.target_wb = wb
                    logger.debug(f"ターゲットファイルを検出: {wb.Name}")
                    break

            # ターゲットワークブックが見つからない場合のエラー
            if self.target_wb is None:
                raise ExcelTransferError(
                    f"ターゲットファイルが開かれていません: {target_basename}\n\n"
                    "Excelで該当ファイルを開いてから実行してください。"
                )

            # ターゲットシートを取得
            try:
                target_sheet_names = [ws.Name for ws in self.target_wb.Worksheets]
                logger.info(f"ターゲットファイルの利用可能なシート: {target_sheet_names}")
                logger.info(f"検索するシート名: '{self.target_sheet}'")

                self.target_ws = self.target_wb.Worksheets(self.target_sheet)
                logger.info(f"✓ ターゲットシートに接続: {self.target_sheet}")
            except Exception as e:
                target_sheet_names = [ws.Name for ws in self.target_wb.Worksheets]
                raise ExcelTransferError(
                    f"ターゲットシートが見つかりません: '{self.target_sheet}'\n\n"
                    f"ファイル: {self.target_filename}\n"
                    f"利用可能なシート: {target_sheet_names}",
                    operation="Excel接続",
                    original_error=e
                ) from e

            # 接続状態の検証
            if self.excel is None or self.target_ws is None:
                raise ExcelTransferError(
                    "Excelへの接続に失敗しました。\n"
                    "ワークシートまたはアプリケーションオブジェクトが取得できませんでした。",
                    operation="Excel接続"
                )

            # データ存在確認（サンプル値をログ出力）
            logger.info("--- データ存在確認 ---")
            sample_d8 = self.target_ws.Range("D8").Value
            sample_c55 = self.target_ws.Range("C55").Value
            sample_c67 = self.target_ws.Range("C67").Value
            logger.info(f"ターゲットシートのサンプル値: D8='{sample_d8}', C55='{sample_c55}', C67='{sample_c67}'")

        except ExcelTransferError:
            raise
        except Exception as e:
            raise ExcelTransferError(
                f"Excelへの接続に失敗しました: {e}",
                operation="Excel接続",
                original_error=e
            ) from e

    def _cleanup_excel(self) -> None:
        """
        Excel COMオブジェクトをクリーンアップ

        Note:
            参照ファイルとターゲットファイルの両方のCOMオブジェクトを解放
        """
        logger.debug("Excel COMオブジェクトをクリーンアップ中...")

        # ワークシート参照を解放
        self.ref_ws = None
        self.target_ws = None

        # ワークブック参照を解放
        self.ref_wb = None
        self.target_wb = None

        # Excelアプリケーション参照を解放
        if self.excel is not None:
            try:
                del self.excel
            except Exception as e:
                logger.warning(f"Excel COMオブジェクト解放エラー: {e}")
            self.excel = None

        # COM終了処理
        if self._com_initialized:
            try:
                pythoncom.CoUninitialize()
                logger.debug("COM終了処理完了")
            except Exception as e:
                logger.warning(f"COM終了処理エラー: {e}")
            self._com_initialized = False

    def execute(self) -> None:
        """
        Excel転記処理を実行

        Raises:
            ExcelTransferError: 転記処理中にエラーが発生した場合
            SystemExit: システム終了が要求された場合
            KeyboardInterrupt: キーボード割り込みが発生した場合
        """
        logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)
        logger.info("Excel自動転記処理を開始")
        logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)

        try:
            # Excelに接続
            self._report_progress("Excelファイルに接続中...")
            self._connect_to_excel()
            logger.info("Excelファイルに接続しました")

            # 3つのループを実行
            self._execute_transfer_loops()

            # 変更を保存
            self._save_target_workbook()

            logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)
            logger.info("Excel自動転記処理が完了しました")
            logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)
            self._report_progress("Excel転記処理が完了しました")

        except ExcelTransferError:
            raise
        except (SystemExit, KeyboardInterrupt):
            logger.warning("処理が中断されました")
            raise
        except Exception as e:
            logger.error(f"Excel転記処理中にエラーが発生: {e}", exc_info=True)
            raise ExcelTransferError(
                f"Excel転記処理中にエラーが発生しました\n"
                f"参照: {self.ref_filename} ({self.ref_sheet})\n"
                f"対象: {self.target_filename} ({self.target_sheet})\n"
                f"エラー: {e}",
                operation="転記処理",
                original_error=e
            ) from e
        finally:
            # 必ずクリーンアップを実行
            self._cleanup_excel()

    def populate_event_names(
        self,
        school_events: List[str],
        student_council_events: List[str],
        other_activities: List[str]
    ) -> Dict[str, int]:
        """
        行事名をターゲットExcelファイルに設定

        Args:
            school_events: 学校行事名リスト（D8~D50に設定）
            student_council_events: 児童会行事名リスト（C55~C62に設定）
            other_activities: その他の教育活動名リスト（C67~C96に設定）

        Returns:
            Dict[str, int]: カテゴリごとの設定件数
            例: {"school_events": 27, "student_council_events": 3, "other_activities": 7}

        Raises:
            ExcelTransferError: 設定処理中にエラーが発生した場合
        """
        logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)
        logger.info("行事名をExcelに設定中...")
        logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)

        try:
            # ターゲットファイルのみに接続（参照ファイル不要）
            self._connect_to_target_only()
            logger.info("Excelファイルに接続しました")

            counts = {"school_events": 0, "student_council_events": 0, "other_activities": 0}

            # カテゴリ1: 学校行事名 (D8~D50)
            logger.info("学校行事名を設定中...")
            for i, event_name in enumerate(school_events):
                row = 8 + i
                if row > 50:
                    logger.warning(f"学校行事名が多すぎます（最大43件）: {len(school_events)}件")
                    break
                if event_name:  # 空文字列をスキップ
                    self.target_ws.Range(f"D{row}").Value = event_name
                    counts["school_events"] += 1
            logger.info(f"学校行事名を{counts['school_events']}件設定しました")

            # カテゴリ2: 児童会行事名 (C55~C62)
            logger.info("児童会行事名を設定中...")
            for i, event_name in enumerate(student_council_events):
                row = 55 + i
                if row > 62:
                    logger.warning(f"児童会行事名が多すぎます（最大8件）: {len(student_council_events)}件")
                    break
                if event_name:  # 空文字列をスキップ
                    self.target_ws.Range(f"C{row}").Value = event_name
                    counts["student_council_events"] += 1
            logger.info(f"児童会行事名を{counts['student_council_events']}件設定しました")

            # カテゴリ3: その他の教育活動名 (C67~C96)
            logger.info("その他の教育活動名を設定中...")
            for i, event_name in enumerate(other_activities):
                row = 67 + i
                if row > 96:
                    logger.warning(f"その他の教育活動名が多すぎます（最大30件）: {len(other_activities)}件")
                    break
                if event_name:  # 空文字列をスキップ
                    self.target_ws.Range(f"C{row}").Value = event_name
                    counts["other_activities"] += 1
            logger.info(f"その他の教育活動名を{counts['other_activities']}件設定しました")

            # 変更を保存
            self._save_target_workbook()

            logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)
            logger.info("行事名の設定が完了しました")
            logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)

            return counts

        except ExcelTransferError:
            raise
        except Exception as e:
            logger.error(f"行事名設定中にエラーが発生: {e}", exc_info=True)
            raise ExcelTransferError(
                f"行事名の設定中にエラーが発生しました\n"
                f"対象: {self.target_filename} ({self.target_sheet})\n"
                f"エラー: {e}",
                operation="行事名設定",
                original_error=e
            ) from e
        finally:
            # 必ずクリーンアップを実行
            self._cleanup_excel()

    def read_event_names_from_excel(self) -> Dict[str, List[str]]:
        """
        ターゲットExcelファイルから行事名を一括読み込み

        Returns:
            Dict[str, List[str]]: カテゴリごとの行事名リスト
            {
                "school_events": [...],
                "student_council_events": [...],
                "other_activities": [...]
            }

        Raises:
            ExcelTransferError: 読み込み処理中にエラーが発生した場合
        """
        logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)
        logger.info("Excelから行事名を読み込み中...")
        logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)

        try:
            # ターゲットファイルのみに接続（参照ファイル不要）
            self._connect_to_target_only()
            logger.info("Excelファイルに接続しました")

            # カテゴリ1: 学校行事名 (D8~D50) を一括取得
            school_range = self.target_ws.Range("D8:D50").Value
            school_events = self._clean_event_names(school_range)
            logger.info(f"学校行事名を{len(school_events)}件読み込みました")

            # カテゴリ2: 児童会行事名 (C55~C62) を一括取得
            council_range = self.target_ws.Range("C55:C62").Value
            student_council_events = self._clean_event_names(council_range)
            logger.info(f"児童会行事名を{len(student_council_events)}件読み込みました")

            # カテゴリ3: その他の教育活動名 (C67~C96) を一括取得
            other_range = self.target_ws.Range("C67:C96").Value
            other_activities = self._clean_event_names(other_range)
            logger.info(f"その他の教育活動名を{len(other_activities)}件読み込みました")

            # データが全て空の場合はエラー
            total = len(school_events) + len(student_council_events) + len(other_activities)
            if total == 0:
                raise ExcelTransferError(
                    "Excelファイルに行事名が見つかりませんでした。\n\n"
                    "D8~D50、C55~C62、C67~C96 列にデータが入力されているか確認してください。",
                    operation="行事名読み込み"
                )

            logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)
            logger.info(f"行事名の読み込みが完了しました（合計: {total}件）")
            logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)

            return {
                "school_events": school_events,
                "student_council_events": student_council_events,
                "other_activities": other_activities
            }

        except ExcelTransferError:
            raise
        except Exception as e:
            logger.error(f"行事名読み込み中にエラーが発生: {e}", exc_info=True)
            raise ExcelTransferError(
                f"行事名の読み込み中にエラーが発生しました\n"
                f"対象: {self.target_filename} ({self.target_sheet})\n"
                f"エラー: {e}",
                operation="行事名読み込み",
                original_error=e
            ) from e
        finally:
            # 必ずクリーンアップを実行
            self._cleanup_excel()

    def _clean_event_names(self, range_values: Any) -> List[str]:
        """
        Excelから取得した行事名リストをクリーニング

        Args:
            range_values: Excelから取得した値（タプルまたは単一値）

        Returns:
            List[str]: クリーニング済みの行事名リスト（空白除外、重複除去、順序維持）
        """
        # タプルをリストに変換
        if range_values is None:
            return []

        if isinstance(range_values, tuple):
            # 2次元タプル: ((val1,), (val2,), ...) → [val1, val2, ...]
            values = [row[0] if isinstance(row, tuple) else row for row in range_values]
        else:
            # 単一値
            values = [range_values]

        # データクリーニング
        cleaned = []
        seen = set()  # 重複チェック用

        for v in values:
            if v is None:
                continue

            # 文字列化して前後の空白を削除、改行文字も除去
            event_name = str(v).strip().replace('\n', '').replace('\r', '').replace('\t', ' ')

            # 空文字列をスキップ
            if not event_name:
                continue

            # 重複をスキップ（大文字小文字区別、順序維持）
            if event_name in seen:
                continue

            seen.add(event_name)
            cleaned.append(event_name)

        return cleaned


def main() -> None:
    """
    スタンドアロン実行用のメイン処理

    Note:
        v3.5.0以降、ファイルパスはGUI側で管理されるため、
        この関数はテスト目的でのみ使用してください。
        実際のファイルパスを引数として指定する必要があります。
    """
    import sys
    from config_loader import ConfigLoader

    if len(sys.argv) < 3:
        print("使用方法: python update_excel_files.py <参照ファイルパス> <対象ファイルパス>")
        print("例: python update_excel_files.py \"C:\\Downloads\\編集用.xlsx\" \"C:\\Desktop\\様式4.xlsx\"")
        sys.exit(1)

    ref_filename = sys.argv[1]
    target_filename = sys.argv[2]

    # 設定ファイルからシート名のみ取得
    config = ConfigLoader()
    ref_sheet = config.get('files', 'excel_reference_sheet')
    target_sheet = config.get('files', 'excel_target_sheet')

    logger.info(f"参照ファイル: {ref_filename}")
    logger.info(f"反映ファイル: {target_filename}")
    logger.info(f"参照シート: {ref_sheet}")
    logger.info(f"反映シート: {target_sheet}")

    transfer = ExcelTransfer(ref_filename, target_filename, ref_sheet, target_sheet)
    transfer.execute()


if __name__ == "__main__":
    # ログ設定
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )

    main()
