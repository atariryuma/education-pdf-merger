"""
Excel自動転記モジュール

年間行事計画（参照ファイル）から様式ファイル（反映ファイル）へ
日付、行事時数、欠時数を自動転記・集計する
"""
import logging
import os
import re
from difflib import SequenceMatcher
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

from exceptions import PDFMergeError, CancelledError
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
            CancelledError: キャンセルされた場合
        """
        if self.cancel_check and self.cancel_check():
            logger.info("ユーザーによって処理がキャンセルされました")
            raise CancelledError("ユーザーによって処理がキャンセルされました")

    def _read_data_row(self, found_row: int, start_col: str, end_col: str) -> list:
        """
        参照Excelから行データを読み取る（キャッシュ優先）

        Args:
            found_row: 行番号
            start_col: 開始列（E～AN用）
            end_col: 終了列

        Returns:
            list: 行データ（文字列に変換済み）
        """
        self._ensure_ref_cache()
        # E～AN列のキャッシュから取得
        if (start_col == ExcelTransferConstants.REF_DATA_START_COL
                and end_col == ExcelTransferConstants.REF_DATA_END_COL):
            return self._ref_data_cache.get(found_row, [])

        # キャッシュ外の範囲はCOMで取得（フォールバック）
        range_str = f"{start_col}{found_row}:{end_col}{found_row}"
        rng = self.ref_ws.Range(range_str).Value
        if rng is None:
            return []
        if isinstance(rng, tuple):
            return [str(cell) if cell is not None else "" for cell in rng[0]]
        return [str(rng) if rng is not None else ""]

    def _normalize_text(self, text: str) -> str:
        """
        テキストを正規化（空白統一、前後トリム）

        Args:
            text: 正規化対象のテキスト

        Returns:
            str: 正規化されたテキスト
        """
        if not text:
            return ""
        return re.sub(r'[\s　]+', ' ', str(text).strip())

    def _get_ref_column_data(self) -> List[Tuple[int, str]]:
        """
        参照ExcelのC列データをキャッシュ付きで一括取得（非空行のみ）

        Returns:
            List[Tuple[int, str]]: (行番号, セル値)のリスト
        """
        self._ensure_ref_cache()
        return self._ref_c_cache

    def _ensure_ref_cache(self) -> None:
        """参照Excelのキャッシュを初期化（C列全行マップ、A列日付、E～ANデータ）"""
        if hasattr(self, '_ref_c_cache'):
            return

        used_range = self.ref_ws.UsedRange
        last_row = used_range.Row + used_range.Rows.Count - 1

        # C列を一括取得（非空リスト + 全行マップの両方を作成）
        range_values = self.ref_ws.Range(f"C1:C{last_row}").Value
        self._ref_c_cache: List[Tuple[int, str]] = []
        self._ref_c_all: Dict[int, str] = {}  # 行番号→値（空行はキーなし）

        if range_values:
            if isinstance(range_values, tuple):
                for i, row in enumerate(range_values):
                    val = row[0] if isinstance(row, tuple) else row
                    if val is not None and str(val).strip():
                        self._ref_c_cache.append((i + 1, str(val)))
                        self._ref_c_all[i + 1] = str(val)
            else:
                if range_values is not None:
                    self._ref_c_cache.append((1, str(range_values)))
                    self._ref_c_all[1] = str(range_values)

        # A列（日付）を一括取得
        a_values = self.ref_ws.Range(f"A1:A{last_row}").Value
        self._ref_a_cache: Dict[int, Any] = {}
        if a_values:
            if isinstance(a_values, tuple):
                for i, row in enumerate(a_values):
                    val = row[0] if isinstance(row, tuple) else row
                    if val is not None:
                        self._ref_a_cache[i + 1] = val
            else:
                if a_values is not None:
                    self._ref_a_cache[1] = a_values

        # E～AN列を一括取得（全行）
        data_range = self.ref_ws.Range(
            f"{ExcelTransferConstants.REF_DATA_START_COL}1:"
            f"{ExcelTransferConstants.REF_DATA_END_COL}{last_row}"
        ).Value
        self._ref_data_cache: Dict[int, list] = {}
        if data_range and isinstance(data_range, tuple):
            for i, row in enumerate(data_range):
                if row and isinstance(row, tuple):
                    row_data = [str(cell) if cell is not None else "" for cell in row]
                    if any(cell for cell in row_data):  # 全空行は除外
                        self._ref_data_cache[i + 1] = row_data

        self._ref_last_row = last_row
        logger.info(
            f"参照Excelキャッシュ完了: C列={len(self._ref_c_cache)}件, "
            f"データ行={len(self._ref_data_cache)}件, 最終行={last_row}"
        )

    def _split_cell_lines(self, cell_value: str) -> List[str]:
        """
        セル内テキストを行・区切り文字で分割

        Args:
            cell_value: セルの値

        Returns:
            List[str]: 分割された行リスト（空行除外）
        """
        # 改行、「・」「※」で分割
        lines = re.split(r'[\n\r]+|(?=[・※])', cell_value)
        return [line.strip().lstrip('・※') for line in lines if line.strip()]

    def _find_value_in_source(self, search_value: str) -> Optional[int]:
        """
        参照ExcelのC列から値を検索（部分一致→あいまい検索の2段階）

        1. 部分文字列一致: 検索値がセル内に含まれるか（複数行セル対応）
        2. あいまい検索: セル内の各行と類似度を比較し、最も近いものを返す

        Args:
            search_value: 検索値

        Returns:
            Optional[int]: 見つかった行番号（見つからない場合None）
        """
        SIMILARITY_THRESHOLD = 0.5

        normalized_search = self._normalize_text(search_value)
        if not normalized_search:
            return None

        ref_data = self._get_ref_column_data()

        # ステップ1: 部分文字列一致（最も短いマッチ＝最も近いものを優先）
        best_partial_row: Optional[int] = None
        best_partial_len: int = float('inf')  # type: ignore[assignment]
        best_partial_value: str = ""

        for row_num, cell_value in ref_data:
            normalized_cell = self._normalize_text(cell_value)
            if normalized_search in normalized_cell:
                if len(normalized_cell) < best_partial_len:
                    best_partial_row = row_num
                    best_partial_len = len(normalized_cell)
                    best_partial_value = cell_value

        if best_partial_row is not None:
            if best_partial_len == len(normalized_search):
                logger.info(
                    f"    完全一致: '{search_value}' = '{best_partial_value}' (行{best_partial_row})"
                )
            else:
                display = best_partial_value[:50] + '...' if len(best_partial_value) > 50 else best_partial_value
                logger.info(
                    f"    部分一致: '{search_value}' ⊂ '{display}' (行{best_partial_row})"
                )
            return best_partial_row

        # ステップ2: あいまい検索（セル内の各行と比較）
        best_row: Optional[int] = None
        best_ratio: float = 0.0
        best_value: str = ""

        for row_num, cell_value in ref_data:
            # セル全体との比較
            normalized_cell = self._normalize_text(cell_value)
            if not normalized_cell:
                continue

            ratio = SequenceMatcher(None, normalized_search, normalized_cell).ratio()
            if ratio > best_ratio:
                best_ratio = ratio
                best_row = row_num
                best_value = cell_value

            # セル内の各行との比較（複数行セル対応）
            for line in self._split_cell_lines(cell_value):
                normalized_line = self._normalize_text(line)
                if not normalized_line:
                    continue

                line_ratio = SequenceMatcher(None, normalized_search, normalized_line).ratio()
                if line_ratio > best_ratio:
                    best_ratio = line_ratio
                    best_row = row_num
                    best_value = line

        if best_ratio >= SIMILARITY_THRESHOLD:
            logger.info(
                f"    あいまい一致: '{search_value}' ≈ '{best_value}' "
                f"(類似度: {best_ratio:.0%}, 行{best_row})"
            )
            return best_row
        else:
            if best_value:
                logger.info(
                    f"    一致なし: '{search_value}' "
                    f"(最も近い: '{best_value}', 類似度: {best_ratio:.0%} < 閾値{SIMILARITY_THRESHOLD:.0%})"
                )
            else:
                logger.info(f"    一致なし: '{search_value}' (参照データなし)")
            return None

    def _read_cell_value(self, row: int, col: str) -> Any:
        """
        参照Excelからセル値を読み取る（キャッシュ優先）

        Args:
            row: 行番号
            col: 列（"A" or "C"）

        Returns:
            Any: セル値
        """
        self._ensure_ref_cache()
        col_upper = col.upper() if isinstance(col, str) else col
        if col_upper == "A":
            return self._ref_a_cache.get(row)
        if col_upper == "C":
            return self._ref_c_all.get(row)
        # その他の列はCOM呼び出し（通常は到達しない）
        if isinstance(col, str) and len(col) == 1:
            col_num = ord(col_upper) - ord('A') + 1
            return self.ref_ws.Cells(row, col_num).Value
        return self.ref_ws.Range(f"{col}{row}").Value

    def _get_period_base_name(self, text: str) -> Optional[str]:
        """
        「～期間」「～週間」のベース名を抽出（連番①②等を除去）

        複数行セルの場合、「期間」「週間」を含む行だけを対象にベース名を抽出する。

        Args:
            text: セルのテキスト

        Returns:
            Optional[str]: ベース名（「期間」「週間」を含まない場合はNone）
        """
        PERIOD_KEYWORDS = ("期間", "週間")
        if not any(kw in text for kw in PERIOD_KEYWORDS):
            return None
        # 複数行セル対応：「期間」「週間」を含む行だけを抽出
        for line in re.split(r'[\n\r]+|(?=[・※])', text):
            line = line.strip().lstrip('・※')
            if not line:
                continue
            if any(kw in line for kw in PERIOD_KEYWORDS):
                # 末尾の日付範囲サフィックス（～24日、～10月9日）を除去
                line = re.sub(r'～\d{1,2}月?\d{1,2}日$', '', line).strip()
                # 末尾の連番記号（①②③...、(1)(2)、１２３...）を除去してベース名を取得
                return re.sub(r'[①-⑳⓪-⓴\(（]\d*[\)）]?|\d+$', '', line).strip()
        return None

    def _parse_date_range_suffix(self, text: str) -> Optional[Tuple[int, int]]:
        """
        検索値から「～XX日」「～X月X日」サフィックスを解析

        Args:
            text: ターゲットExcelの検索値

        Returns:
            Optional[Tuple[int, int]]: (月, 日) のタプル。該当なしはNone
        """
        # ～10月9日 のような月またぎパターン
        m = re.search(r'～(\d{1,2})月(\d{1,2})日', text)
        if m:
            return (int(m.group(1)), int(m.group(2)))
        # ～24日 のような同月パターン（月は起点行のA列から取得）
        m = re.search(r'～(\d{1,2})日', text)
        if m:
            return (0, int(m.group(1)))  # 月=0は「起点行と同月」を意味
        return None

    def _get_period_rows(
        self, found_row: int, search_value: Optional[str] = None
    ) -> List[int]:
        """
        「～期間」「～週間」行の関連行を3パターンで収集

        パターン1（日付範囲）: 検索値に「～24日」「～10月9日」→ 日付範囲内の欠時行を収集
        パターン2（連番）: ①②③④のベース名一致行を収集
        パターン3（単一行）: found_rowのみ

        Args:
            found_row: 参照Excelの行番号
            search_value: ターゲットExcelの検索値（日付範囲の判定に使用）

        Returns:
            List[int]: 対象行番号のリスト（ソート済み）
        """
        self._ensure_ref_cache()
        search_str = str(search_value).strip() if search_value else ""

        # --- パターン1: 日付範囲（～XX日 / ～X月X日） ---
        # まずターゲットの検索値をチェック
        date_suffix = self._parse_date_range_suffix(search_str)
        if date_suffix is not None:
            return self._collect_rows_by_date_range(found_row, date_suffix, search_str)

        # ターゲットになくても、参照ExcelのC列に日付範囲がある場合はパターン1を適用
        c_value = self._read_cell_value(found_row, "C")
        if c_value is not None:
            c_text = str(c_value)
            # C列の「期間」「週間」を含む行から日付範囲を検出
            PERIOD_KEYWORDS = ("期間", "週間")
            if any(kw in c_text for kw in PERIOD_KEYWORDS):
                for line in re.split(r'[\n\r]+|(?=[・※])', c_text):
                    line_stripped = line.strip().lstrip('・※')
                    if any(kw in line_stripped for kw in PERIOD_KEYWORDS):
                        ref_date_suffix = self._parse_date_range_suffix(line_stripped)
                        if ref_date_suffix is not None:
                            logger.info(
                                f"    期間収集: '{search_str}' → 参照C列から日付範囲検出 "
                                f"('{line_stripped}')"
                            )
                            return self._collect_rows_by_date_range(
                                found_row, ref_date_suffix, search_str
                            )

        # --- パターン2: 連番（①②③④のベース名一致） ---
        if c_value is not None:
            base_name = self._get_period_base_name(str(c_value))
            if base_name is not None:
                return self._collect_rows_by_base_name(found_row, base_name, search_str)

        # --- パターン3: 単一行 ---
        logger.info(f"    期間収集: '{search_str}' → パターン3(単一行) 行{found_row}")
        return [found_row]

    def _collect_rows_by_date_range(
        self, found_row: int, date_suffix: Tuple[int, int], search_str: str
    ) -> List[int]:
        """
        パターン1: 日付範囲で関連行を収集

        起点行のA列日付から終了日までの範囲で、欠時セルを含む行を収集する。

        Args:
            found_row: 参照Excelの起点行
            date_suffix: (月, 日) - 月=0は起点行と同月
            search_str: ログ出力用の検索値
        """
        import datetime as dt

        # 起点行のA列から開始日を取得
        start_date_raw = self._ref_a_cache.get(found_row)
        if start_date_raw is None:
            logger.warning(
                f"    期間収集: '{search_str}' → 起点行{found_row}にA列日付なし、単一行で処理"
            )
            return [found_row]

        # datetime変換（タイムゾーン情報は除去して統一）
        if isinstance(start_date_raw, dt.datetime):
            start_date = start_date_raw.replace(tzinfo=None)
        elif isinstance(start_date_raw, (int, float)):
            # Excelシリアル値の場合
            try:
                start_date = dt.datetime(1899, 12, 30) + dt.timedelta(days=int(start_date_raw))
            except (ValueError, OverflowError):
                logger.warning(f"    期間収集: 起点行{found_row}の日付変換失敗: {start_date_raw}")
                return [found_row]
        else:
            logger.warning(f"    期間収集: 起点行{found_row}の日付形式不明: {type(start_date_raw)}")
            return [found_row]

        # 終了日を決定
        end_month, end_day = date_suffix
        if end_month == 0:
            # 同月パターン（～24日）
            end_month = start_date.month
        try:
            end_year = start_date.year
            # 月が起点より小さい場合は年跨ぎ（例: 起点12月、終了1月）
            if end_month < start_date.month:
                end_year += 1
            end_date = dt.datetime(end_year, end_month, end_day)
        except ValueError as e:
            logger.warning(f"    期間収集: 終了日の生成失敗 ({end_month}月{end_day}日): {e}")
            return [found_row]

        logger.info(
            f"    期間収集: '{search_str}' → パターン1(日付範囲) "
            f"{start_date.strftime('%m/%d')}～{end_date.strftime('%m/%d')}"
        )

        # 日付範囲内で欠時を含む行を収集
        rows: List[int] = []
        for row_num, date_val in self._ref_a_cache.items():
            # datetime変換（タイムゾーン情報は除去して統一）
            if isinstance(date_val, dt.datetime):
                row_date = date_val.replace(tzinfo=None)
            elif isinstance(date_val, (int, float)):
                try:
                    row_date = dt.datetime(1899, 12, 30) + dt.timedelta(days=int(date_val))
                except (ValueError, OverflowError):
                    continue
            else:
                continue

            # 日付範囲チェック
            if start_date <= row_date <= end_date:
                # 欠時セルがあるか確認
                if row_num in self._ref_data_cache:
                    row_data = self._ref_data_cache[row_num]
                    has_absent = any(
                        str(cell).strip() == ExcelTransferConstants.ABSENT_KEYWORD
                        for cell in row_data if cell is not None
                    )
                    if has_absent:
                        rows.append(row_num)

        rows = sorted(set(rows))
        logger.info(
            f"    → {len(rows)}行を収集 "
            f"(行番号: {rows if len(rows) <= 10 else str(rows[:10]) + '...'})"
        )
        return rows if rows else [found_row]

    def _collect_rows_by_base_name(
        self, found_row: int, base_name: str, search_str: str
    ) -> List[int]:
        """
        パターン2: ベース名一致で連番行を収集（①②③④等）

        Args:
            found_row: 参照Excelの起点行
            base_name: 期間/週間のベース名
            search_str: ログ出力用の検索値
        """
        rows_set: set = {found_row}

        for row_num, cell_value in self._ref_c_cache:
            row_base = self._get_period_base_name(str(cell_value))
            if row_base == base_name:
                rows_set.add(row_num)

        rows = sorted(rows_set)
        logger.info(
            f"    期間収集: '{search_str}' → パターン2(連番) ベース名='{base_name}' "
            f"{len(rows)}行を収集 (行番号: {rows})"
        )
        return rows

    def _detect_event_category(self, found_row: int) -> str:
        """
        参照ExcelのE～AN列から行事カテゴリを自動検出

        Args:
            found_row: 参照Excelの行番号

        Returns:
            str: 検出されたカテゴリ（儀式、文化等）。見つからない場合は空文字
        """
        row_data = self._read_data_row(
            found_row,
            ExcelTransferConstants.REF_DATA_START_COL,
            ExcelTransferConstants.REF_DATA_END_COL
        )
        if not row_data:
            return ""

        # 各EVENT_KEYWORDの出現回数をカウント
        keyword_counts: Dict[str, int] = {}
        for cell in row_data:
            if cell is None:
                continue
            cell_str = str(cell).strip()
            if not cell_str:
                continue
            for kw in ExcelTransferConstants.EVENT_KEYWORDS:
                if kw in cell_str:
                    keyword_counts[kw] = keyword_counts.get(kw, 0) + 1
                    break

        if not keyword_counts:
            # EVENT_KEYWORDSに一致しない場合、欠時チェック
            for cell in row_data:
                if cell is not None and str(cell).strip() == ExcelTransferConstants.ABSENT_KEYWORD:
                    return ExcelTransferConstants.ABSENT_KEYWORD
            return ""

        # 最も多いキーワードを返す
        return max(keyword_counts, key=keyword_counts.get)  # type: ignore[arg-type]

    def _count_events_in_found_row(
        self,
        found_row: int,
        filter_keyword: Optional[str] = None,
        search_value: Optional[str] = None
    ) -> Dict[int, Tuple[int, int]]:
        """
        参照Excelの found_row の【E～AN】列を一括で取得し、学年毎にカウント

        「～期間」「～週間」行の場合は、関連行も含めて合算する。

        Args:
            found_row: 参照Excelの行番号
            filter_keyword: フィルター用キーワード（D8～D50用、Noneの場合はループ2,3）
            search_value: ターゲットの検索値（期間/週間の判定に使用）

        Returns:
            辞書 {grade: (event_count, absent_count), ...}（grade 1～6）

        Raises:
            SystemExit, KeyboardInterrupt: 即座に再スロー
        """
        try:
            # ターゲットの検索値から期間/週間を判定（参照C列ではなく検索値で判定）
            search_str = str(search_value).strip() if search_value else ""
            is_period = any(kw in search_str for kw in ("期間", "週間"))

            # 期間/週間の場合は関連行を全て収集
            if is_period:
                target_rows = self._get_period_rows(found_row, search_value)
            else:
                target_rows = [found_row]

            # 全対象行のデータを収集
            all_row_data: List[list] = []
            for row in target_rows:
                row_data = self._read_data_row(
                    row,
                    ExcelTransferConstants.REF_DATA_START_COL,
                    ExcelTransferConstants.REF_DATA_END_COL
                )
                if row_data:
                    all_row_data.append(row_data)

            if not all_row_data:
                logger.warning(f"行 {found_row} のデータが空です")
                return {
                    grade: (0, 0)
                    for grade in range(1, ExcelTransferConstants.GRADES_COUNT + 1)
                }

            # 期間/週間の種別判定：どちらも欠時としてカウント
            period_mode: Optional[str] = None
            if is_period:
                period_mode = "absent"
                logger.info(f"    期間/週間検出: '{search_str}' → 欠時としてカウント")

            counts: Dict[int, Tuple[int, int]] = {}

            for grade in range(1, ExcelTransferConstants.GRADES_COUNT + 1):
                start_index = (grade - 1) * ExcelTransferConstants.PERIODS_PER_GRADE
                end_index = start_index + ExcelTransferConstants.PERIODS_PER_GRADE

                total_event = 0
                total_absent = 0

                for row_data in all_row_data:
                    # データ長の検証
                    if end_index > len(row_data):
                        continue

                    group = row_data[start_index:end_index]

                    if is_period:
                        # 期間/週間グループ：「欠時」セルのみカウント
                        # （同日に他の行事キーワードが混在していても欠時だけ拾う）
                        for cell in group:
                            if cell is None:
                                continue
                            cell_str = str(cell).strip()
                            if cell_str == ExcelTransferConstants.ABSENT_KEYWORD:
                                total_absent += 1
                    elif filter_keyword is not None:
                        # D8～D50用：行事キーワードのみカウント（欠時は期間/週間でカウント）
                        for cell in group:
                            if cell is None:
                                continue
                            cell_str = str(cell).strip()
                            if not cell_str:
                                continue
                            if cell_str == filter_keyword:
                                total_event += 1
                    else:
                        # ループ2,3用（通常行事）：行事時数のみカウント
                        total_event += sum(
                            1 for cell in group
                            if cell is not None and any(
                                keyword in str(cell)
                                for keyword in ExcelTransferConstants.EVENT_KEYWORDS
                            )
                        )

                counts[grade] = (total_event, total_absent)

            # 計算過程をログ出力
            if is_period:
                grade_details = ", ".join(
                    f"{g}年={counts[g][1]}欠時" for g in range(1, ExcelTransferConstants.GRADES_COUNT + 1)
                )
                logger.info(
                    f"    計算結果: '{search_str}' ({len(target_rows)}行合算) → {grade_details}"
                )
            else:
                grade_details = ", ".join(
                    f"{g}年={counts[g][0]}行事/{counts[g][1]}欠時"
                    for g in range(1, ExcelTransferConstants.GRADES_COUNT + 1)
                    if counts[g][0] > 0 or counts[g][1] > 0
                )
                if grade_details:
                    logger.info(f"    計算結果: '{search_str}' → {grade_details}")

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

        # 期間/週間の連番処理
        search_str = str(search_value).strip()
        # original_search_value: 日付範囲サフィックス付きの元値（_get_period_rowsに渡す）
        original_search_value = search_str
        if any(kw in search_str for kw in ("期間", "週間")):
            if re.search(r'[②-⑳]', search_str):
                # ②以降 → A～P列を一括クリア（①に集約）
                empty_row = [[""] * 16]
                self.target_ws.Range(f"A{row}:P{row}").Value = empty_row
                logger.info(f"  - 行{row}: '{search_value}' → ①に集約のためクリア")
                return
            # ① → 連番サフィックスを除去して名前を整理
            clean_name = re.sub(r'[①-⑳]', '', search_str).strip()
            if clean_name != search_str:
                self.target_ws.Range(f"{search_col}{row}").Value = clean_name
                logger.info(f"  - 行{row}: '{search_str}' → '{clean_name}' に整理")
                search_value = clean_name
                original_search_value = clean_name

            # ～XX日 / ～X月X日 サフィックスを検索値から除去（検索用）
            # 元の値はoriginal_search_valueに保持（日付範囲パターンで使用）
            search_for_lookup = re.sub(r'～\d{1,2}月?\d{1,2}日', '', str(search_value)).strip()
            if search_for_lookup != str(search_value):
                logger.info(
                    f"  - 行{row}: 検索用に '～' サフィックス除去: "
                    f"'{search_value}' → '{search_for_lookup}'"
                )
                search_value = search_for_lookup

        logger.info(f"  > 行{row}: {search_col}列の値 '{search_value}' を参照Excelで検索中...")

        # 参照Excelから検索
        found_row = self._find_value_in_source(search_value)

        if found_row is not None:
            # 日付を取得（キャッシュから）
            ref_date = self._read_cell_value(
                found_row,
                ExcelTransferConstants.TARGET_DATE_COL
            )

            # Loop1（D列検索）の場合、C列（内容）を参照データから自動検出
            if search_col == ExcelTransferConstants.LOOP1_SEARCH_COL:
                detected_category = self._detect_event_category(found_row)
                if detected_category:
                    self.target_ws.Range(
                        f"{ExcelTransferConstants.TARGET_FILTER_COL}{row}"
                    ).Value = detected_category
                    filter_keyword = detected_category
                    logger.info(f"    内容自動検出: '{detected_category}'")

            # 行事時数・欠時数をカウント（日付範囲サフィックス付きの元値を渡す）
            counts = self._count_events_in_found_row(found_row, filter_keyword, original_search_value)

            # 全学年で行事時数・欠時数ともに0の場合は転記から除外
            has_any_count = any(
                event_count > 0 or absent_count > 0
                for event_count, absent_count in counts.values()
            )
            if not has_any_count:
                # カウントなし → 行全体をクリア（ソートで下に移動）
                self.target_ws.Range(f"A{row}:P{row}").Value = [[""] * 16]
                logger.info(
                    f"  - 行{row}: '{search_value}' → 参照あり(行{found_row})だが時数なし、除外"
                )
                return

            # A～P列の書き込みデータを構築
            self.target_ws.Range(f"A{row}").Value = ref_date

            # E～P列を一括書き込み（12セル = 6学年 × 2列）
            grade_data = []
            for grade in range(1, ExcelTransferConstants.GRADES_COUNT + 1):
                event_count, absent_count = counts[grade]
                grade_data.append(event_count if event_count else "")
                grade_data.append(absent_count if absent_count else "")
            self.target_ws.Range(f"E{row}:P{row}").Value = [grade_data]

            logger.info(
                f"  ✓ 行{row}: '{search_value}' → 参照Excel行{found_row} (日付: {ref_date})"
            )
        else:
            # 見つからない場合：行全体をクリア（ソートで下に移動）
            self.target_ws.Range(f"A{row}:P{row}").Value = [[""] * 16]

            logger.warning(
                f"  ✗ 行{row}: '{search_value}' → 参照Excelに該当なし、除外"
            )

    def _collect_merge_areas(self, range_str: str) -> List[str]:
        """
        指定範囲内の結合セルアドレスを収集

        Args:
            range_str: 対象範囲（例: "A8:P50"）

        Returns:
            List[str]: 結合セルのアドレスリスト（例: ["C8:D8", "C9:E9"]）
        """
        merge_addresses: List[str] = []
        seen: set = set()

        target_range = self.target_ws.Range(range_str)
        for cell in target_range:
            if cell.MergeCells:
                addr = cell.MergeArea.Address
                if addr not in seen:
                    seen.add(addr)
                    merge_addresses.append(addr)

        logger.debug(f"結合セルを{len(merge_addresses)}件検出: {merge_addresses}")
        return merge_addresses

    def _restore_merge_areas(self, merge_addresses: List[str]) -> None:
        """
        結合セルを復元

        Args:
            merge_addresses: 結合セルのアドレスリスト
        """
        for addr in merge_addresses:
            try:
                self.target_ws.Range(addr).Merge()
            except Exception as e:
                logger.warning(f"結合セルの復元に失敗 ({addr}): {e}")

        logger.debug(f"結合セルを{len(merge_addresses)}件復元しました")

    def _sort_range(
        self, range_str: str, key_cell: str, key_cell2: Optional[str] = None
    ) -> None:
        """
        指定範囲を昇順ソート（結合セルを保持）

        結合セルがある場合、結合アドレスを記録→解除→ソート→再結合する。

        Args:
            range_str: ソート範囲（例: "A8:P50"）
            key_cell: 第1ソートキー（例: "C8"）
            key_cell2: 第2ソートキー（例: "A8"）。Noneの場合は単一キー

        Raises:
            ExcelTransferError: 並び替えに失敗した場合
        """
        try:
            # 結合セルのアドレスを記録
            merge_addresses = self._collect_merge_areas(range_str)

            if merge_addresses:
                # 結合を解除してからソート
                self.target_ws.Range(range_str).UnMerge()
                logger.info(f"ソート前に結合セルを{len(merge_addresses)}件解除しました")

            # ソート実行
            sort_range = self.target_ws.Range(range_str)
            if key_cell2:
                sort_range.Sort(
                    Key1=self.target_ws.Range(key_cell),
                    Order1=ExcelSortOrder.ASCENDING,
                    Key2=self.target_ws.Range(key_cell2),
                    Order2=ExcelSortOrder.ASCENDING,
                    Header=ExcelSortHeader.NO
                )
            else:
                sort_range.Sort(
                    Key1=self.target_ws.Range(key_cell),
                    Order1=ExcelSortOrder.ASCENDING,
                    Header=ExcelSortHeader.NO
                )

            if merge_addresses:
                # 結合セルを復元
                self._restore_merge_areas(merge_addresses)
                logger.info(f"範囲 {range_str} を並び替え、結合セルを復元しました")
            else:
                logger.info(f"範囲 {range_str} を並び替えました")

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

        total_rows = end_row - start_row
        for i, row in enumerate(range(start_row, end_row)):
            filter_keyword = filter_list[i] if i < len(filter_list) else None
            if filter_keyword is not None:
                filter_keyword = str(filter_keyword).strip()

            self._report_progress(f"ループ1: 転記中... ({i + 1}/{total_rows})")
            self._process_row(row, ExcelTransferConstants.LOOP1_SEARCH_COL, filter_keyword)

        # ループ1の範囲を並び替え（C列=内容ごと → A列=日付順）
        self._report_progress("ループ1: 並び替え中...")
        self._sort_range(
            ExcelTransferConstants.LOOP1_SORT_RANGE,
            ExcelTransferConstants.LOOP1_SORT_KEY,
            "A8"
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
        total_rows = end_row - start_row
        logger.info(f"ループ2: 行{start_row}～{end_row - 1}を処理します（全{total_rows}行）")

        for i, row in enumerate(range(start_row, end_row)):
            self._report_progress(f"ループ2: 転記中... ({i + 1}/{total_rows})")
            self._process_row(row, ExcelTransferConstants.LOOP2_SEARCH_COL, None)

        # ループ2の範囲を並び替え
        self._report_progress("ループ2: 並び替え中...")
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
        total_rows = end_row - start_row
        logger.info(f"ループ3: 行{start_row}～{end_row - 1}を処理します（全{total_rows}行）")

        for i, row in enumerate(range(start_row, end_row)):
            self._report_progress(f"ループ3: 転記中... ({i + 1}/{total_rows})")
            self._process_row(row, ExcelTransferConstants.LOOP3_SEARCH_COL, None)

        # ループ3の範囲を並び替え
        self._sort_range(
            ExcelTransferConstants.LOOP3_SORT_RANGE,
            ExcelTransferConstants.LOOP3_SORT_KEY
        )
        logger.info("【ループ3】完了")

    def _init_com_connection(self) -> None:
        """
        COM初期化とExcelインスタンスへの接続

        Raises:
            ExcelTransferError: 接続に失敗した場合
        """
        try:
            pythoncom.CoInitialize()
            self._com_initialized = True
            logger.debug("COM初期化完了")
        except Exception as e:
            logger.debug(f"COM初期化スキップ（既に初期化済み）: {e}")
            self._com_initialized = False

        self.excel = win32com.client.Dispatch("Excel.Application")
        logger.debug("Excelインスタンスに接続しました")

    def _find_workbook(self, filename: str) -> Any:
        """
        開いているワークブックをファイル名で検索

        Args:
            filename: ファイル名（フルパスまたはベース名）

        Returns:
            Any: 見つかったワークブックオブジェクト

        Raises:
            ExcelTransferError: ワークブックが見つからない場合
        """
        basename = os.path.basename(filename)
        for wb in self.excel.Workbooks:
            if basename == wb.Name:
                logger.debug(f"ワークブックを検出: {wb.Name}")
                return wb
        raise ExcelTransferError(
            f"ファイルが開かれていません: {basename}\n\n"
            "Excelで該当ファイルを開いてから実行してください。"
        )

    def _connect_worksheet(self, workbook: Any, sheet_name: str, filename: str) -> Any:
        """
        ワークブックからシートに接続

        Args:
            workbook: ワークブックオブジェクト
            sheet_name: シート名
            filename: ファイル名（エラーメッセージ用）

        Returns:
            Any: ワークシートオブジェクト

        Raises:
            ExcelTransferError: シートが見つからない場合
        """
        try:
            sheet_names = [ws.Name for ws in workbook.Worksheets]
            logger.info(f"利用可能なシート: {sheet_names}")
            logger.info(f"検索するシート名: '{sheet_name}'")

            worksheet = workbook.Worksheets(sheet_name)
            logger.info(f"✓ シートに接続: {sheet_name}")
            return worksheet
        except Exception as e:
            sheet_names = [ws.Name for ws in workbook.Worksheets]
            raise ExcelTransferError(
                f"シートが見つかりません: '{sheet_name}'\n\n"
                f"ファイル: {filename}\n"
                f"利用可能なシート: {sheet_names}",
                operation="Excel接続",
                original_error=e
            ) from e

    def _log_target_sample_data(self) -> None:
        """ターゲットシートのサンプルデータをログ出力"""
        logger.info("--- データ存在確認 ---")
        sample_d8 = self.target_ws.Range("D8").Value
        sample_c55 = self.target_ws.Range("C55").Value
        sample_c67 = self.target_ws.Range("C67").Value
        logger.info(f"サンプル値: D8='{sample_d8}', C55='{sample_c55}', C67='{sample_c67}'")

    def _connect_to_excel(self) -> None:
        """
        既存のExcelインスタンスに接続してワークブック/シートを取得

        Raises:
            ExcelTransferError: 接続に失敗した場合
        """
        try:
            self._init_com_connection()

            # 参照・ターゲット両方のワークブックを検索
            self.ref_wb = self._find_workbook(self.ref_filename)
            self.target_wb = self._find_workbook(self.target_filename)

            # シートに接続
            self.ref_ws = self._connect_worksheet(
                self.ref_wb, self.ref_sheet, self.ref_filename
            )
            self.target_ws = self._connect_worksheet(
                self.target_wb, self.target_sheet, self.target_filename
            )

            self._log_target_sample_data()

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
            self._init_com_connection()

            # ターゲットワークブックのみ検索
            self.target_wb = self._find_workbook(self.target_filename)

            # ターゲットシートに接続
            self.target_ws = self._connect_worksheet(
                self.target_wb, self.target_sheet, self.target_filename
            )

            self._log_target_sample_data()

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

        except (CancelledError, ExcelTransferError):
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
