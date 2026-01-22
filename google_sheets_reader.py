"""
Google Sheetsデータリーダー

Google Sheets APIを使用してスプレッドシートからデータを読み取ります。
"""
import logging
import re
from typing import Any, List, Optional, Tuple

try:
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
    from google.oauth2.credentials import Credentials
except ImportError as e:
    raise ImportError(
        "Google APIライブラリがインストールされていません。\n"
        "以下のコマンドを実行してインストールしてください:\n"
        "pip install google-api-python-client"
    ) from e

from exceptions import GoogleSheetsError

# ロガーの設定
logger = logging.getLogger(__name__)


class GoogleSheetsReader:
    """
    Google Sheetsデータリーダー

    Google Sheets APIを使用してスプレッドシートからデータを読み取ります。
    """

    def __init__(self, credentials: Credentials, spreadsheet_url: str, sheet_name: str) -> None:
        """
        初期化

        Args:
            credentials: Google API認証情報
            spreadsheet_url: スプレッドシートのURL
            sheet_name: シート名

        Raises:
            GoogleSheetsError: URL解析に失敗した場合
        """
        self.credentials = credentials
        self.sheet_name = sheet_name

        # URLからスプレッドシートIDを抽出
        self.spreadsheet_id = self.parse_spreadsheet_url(spreadsheet_url)
        logger.debug(f"スプレッドシートID: {self.spreadsheet_id}")
        logger.debug(f"シート名: {self.sheet_name}")

        # Google Sheets APIサービスを構築
        try:
            self.service = build('sheets', 'v4', credentials=self.credentials)
            logger.debug("Google Sheets APIサービスを構築しました")
        except Exception as e:
            raise GoogleSheetsError(
                "Google Sheets APIサービスの構築に失敗しました",
                operation="初期化",
                original_error=e
            ) from e

    @staticmethod
    def parse_spreadsheet_url(url: str) -> str:
        """
        スプレッドシートURLからIDを抽出

        Args:
            url: スプレッドシートのURL

        Returns:
            スプレッドシートID

        Raises:
            GoogleSheetsError: URL解析に失敗した場合
        """
        # URLパターン: https://docs.google.com/spreadsheets/d/{spreadsheet_id}/...
        pattern = r'/spreadsheets/d/([a-zA-Z0-9-_]+)'
        match = re.search(pattern, url)

        if not match:
            raise GoogleSheetsError(
                f"無効なGoogle Sheets URLです: {url}\n"
                "正しいURL形式: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/...",
                operation="URL解析"
            )

        return match.group(1)

    def read_range(self, range_addr: str) -> List[List[Any]]:
        """
        範囲を読み取り

        Args:
            range_addr: A1形式の範囲指定（例: "E8:AN50"）
                       シート名は自動的に付加されます

        Returns:
            2次元リスト（行×列）
            空セルは空文字列として返されます

        Raises:
            GoogleSheetsError: データ読み取りに失敗した場合
        """
        # シート名を含む範囲指定
        full_range = f"{self.sheet_name}!{range_addr}"

        try:
            logger.debug(f"範囲読み取り: {full_range}")
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range=full_range,
                valueRenderOption='UNFORMATTED_VALUE'  # 数式ではなく値を取得
            ).execute()

            values = result.get('values', [])
            logger.debug(f"読み取り完了: {len(values)}行")

            return values

        except HttpError as e:
            error_reason = e.error_details[0].get('reason', 'UNKNOWN') if e.error_details else 'UNKNOWN'

            if error_reason == 'notFound':
                raise GoogleSheetsError(
                    f"スプレッドシートまたはシートが見つかりません: {full_range}",
                    operation="データ読み取り",
                    original_error=e
                ) from e
            else:
                raise GoogleSheetsError(
                    f"データの読み取りに失敗しました: {full_range}",
                    operation="データ読み取り",
                    original_error=e
                ) from e

        except Exception as e:
            raise GoogleSheetsError(
                f"データの読み取り中に予期しないエラーが発生しました: {full_range}",
                operation="データ読み取り",
                original_error=e
            ) from e

    def read_cell(self, row: int, column: str) -> Any:
        """
        単一セルを読み取り

        Args:
            row: 行番号（1始まり）
            column: 列名（A, B, C, ...）

        Returns:
            セルの値（空の場合は空文字列）

        Raises:
            GoogleSheetsError: データ読み取りに失敗した場合
        """
        cell_addr = f"{column}{row}"
        values = self.read_range(cell_addr)

        if not values or not values[0]:
            return ""

        return values[0][0]

    def find_value(self, search_value: str, search_column: str,
                   start_row: int = 1, end_row: int = 1000) -> Optional[int]:
        """
        列内で値を検索（部分一致）

        Args:
            search_value: 検索する値
            search_column: 検索対象の列（例: "C"）
            start_row: 検索開始行（デフォルト: 1）
            end_row: 検索終了行（デフォルト: 1000）

        Returns:
            見つかった行番号（1始まり）、見つからない場合はNone

        Raises:
            GoogleSheetsError: データ読み取りに失敗した場合
        """
        # 検索範囲を読み取り
        range_addr = f"{search_column}{start_row}:{search_column}{end_row}"
        values = self.read_range(range_addr)

        if not values:
            logger.debug(f"検索範囲にデータがありません: {range_addr}")
            return None

        # 部分一致で検索
        search_value_lower = str(search_value).lower()

        for i, row_data in enumerate(values):
            if not row_data:  # 空行スキップ
                continue

            cell_value = str(row_data[0]).lower()

            if search_value_lower in cell_value:
                found_row = start_row + i
                logger.debug(f"値が見つかりました: '{search_value}' -> 行{found_row}")
                return found_row

        logger.debug(f"値が見つかりませんでした: '{search_value}' in {range_addr}")
        return None

    def read_row_range(self, row: int, start_col: str, end_col: str) -> List[Any]:
        """
        特定の行の範囲を読み取り

        Args:
            row: 行番号（1始まり）
            start_col: 開始列（例: "E"）
            end_col: 終了列（例: "AN"）

        Returns:
            セル値のリスト（空セルは空文字列）

        Raises:
            GoogleSheetsError: データ読み取りに失敗した場合
        """
        range_addr = f"{start_col}{row}:{end_col}{row}"
        values = self.read_range(range_addr)

        if not values or not values[0]:
            # 空の行の場合、列数分の空文字列を返す
            col_count = self._column_distance(start_col, end_col) + 1
            return [""] * col_count

        return values[0]

    @staticmethod
    def _column_distance(start_col: str, end_col: str) -> int:
        """
        2つの列の間の距離を計算

        Args:
            start_col: 開始列（例: "E"）
            end_col: 終了列（例: "AN"）

        Returns:
            列の距離（例: E→AN = 35）
        """
        def col_to_num(col: str) -> int:
            """列名を数値に変換（A=0, B=1, ...）"""
            num = 0
            for char in col:
                num = num * 26 + (ord(char) - ord('A') + 1)
            return num - 1

        return col_to_num(end_col) - col_to_num(start_col)
