"""
Google Sheets → Excel自動転記モジュール

Google Sheets（参照元）からExcelファイル（ターゲット）へ
日付、行事時数、欠時数を自動転記・集計する
"""
import logging
import os
from typing import Any, Optional, Callable

try:
    import win32com.client
    import pythoncom
except ImportError as e:
    raise ImportError(
        "pywin32 がインストールされていません。\n"
        "以下のコマンドを実行してインストールしてください:\n"
        "pip install pywin32"
    ) from e

from google.oauth2.credentials import Credentials

from base_excel_transfer import BaseExcelTransfer
from exceptions import GoogleSheetsError
from constants import PDFConversionConstants
from google_sheets_reader import GoogleSheetsReader

# ロガーの設定
logger = logging.getLogger(__name__)


class GoogleSheetsTransfer(BaseExcelTransfer):
    """Google Sheets → Excel自動転記処理クラス"""

    def __init__(
        self,
        credentials: Credentials,
        sheets_url: str,
        ref_sheet: str,
        target_filename: str,
        target_sheet: str,
        progress_callback: Optional[Callable[[str], None]] = None,
        cancel_check: Optional[Callable[[], bool]] = None
    ) -> None:
        """
        初期化

        Args:
            credentials: Google API認証情報
            sheets_url: 参照元のGoogle Sheets URL
            ref_sheet: 参照シート名
            target_filename: ターゲットExcelファイル名
            target_sheet: ターゲットシート名
            progress_callback: 進捗状況を報告するコールバック関数
            cancel_check: キャンセル確認用コールバック関数（Trueで中断）
        """
        super().__init__(target_filename, target_sheet, progress_callback, cancel_check)

        self.credentials = credentials
        self.sheets_url = sheets_url
        self.ref_sheet = ref_sheet

        # Google Sheetsリーダー
        self.sheets_reader: Optional[GoogleSheetsReader] = None

    def _get_error_class(self) -> type:
        """
        エラークラスを取得

        Returns:
            type: GoogleSheetsError
        """
        return GoogleSheetsError

    def _read_data_row(self, found_row: int, start_col: str, end_col: str) -> list:
        """
        Google Sheetsから行データを読み取る

        Args:
            found_row: 行番号
            start_col: 開始列
            end_col: 終了列

        Returns:
            list: 行データ
        """
        return self.sheets_reader.read_row_range(found_row, start_col, end_col)

    def _find_value_in_source(self, search_value: str) -> Optional[int]:
        """
        Google SheetsのC列から値を検索

        Args:
            search_value: 検索値

        Returns:
            Optional[int]: 見つかった行番号（見つからない場合None）
        """
        return self.sheets_reader.find_value(
            search_value,
            search_column="C",
            start_row=1,
            end_row=1000
        )

    def _read_cell_value(self, row: int, col: str) -> Any:
        """
        Google Sheetsからセル値を読み取る

        Args:
            row: 行番号
            col: 列

        Returns:
            Any: セル値
        """
        return self.sheets_reader.read_cell(row, col)

    def _connect_to_sheets(self) -> None:
        """
        Google Sheetsに接続

        Raises:
            GoogleSheetsError: 接続に失敗した場合
        """
        try:
            logger.info("Google Sheetsに接続中...")
            self.sheets_reader = GoogleSheetsReader(
                credentials=self.credentials,
                spreadsheet_url=self.sheets_url,
                sheet_name=self.ref_sheet
            )
            logger.info("Google Sheetsに接続しました")
        except Exception as e:
            raise GoogleSheetsError(
                "Google Sheetsへの接続に失敗しました",
                operation="接続",
                original_error=e
            ) from e

    def _connect_to_target_excel(self) -> None:
        """
        ターゲットExcelファイルに接続

        既存のExcelアプリケーションから指定されたワークブックとシートを取得します。

        Raises:
            GoogleSheetsError: Excel接続に失敗した場合、またはワークブック/シートが見つからない場合
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

            # ターゲットワークブックを取得（ファイル名で検索）
            self.target_wb = None

            # フルパスの場合はファイル名のみを抽出
            target_basename = os.path.basename(self.target_filename)

            for wb in self.excel.Workbooks:
                if target_basename in wb.Name:
                    self.target_wb = wb
                    logger.debug(f"ターゲットファイルを検出: {wb.Name}")
                    break

            # ワークブックが見つからない場合のエラー
            if self.target_wb is None:
                raise GoogleSheetsError(
                    f"ターゲットファイルが開かれていません: {target_basename}\n\n"
                    "Excelで該当ファイルを開いてから実行してください。",
                    operation="Excel接続"
                )

            # シートを取得
            try:
                self.target_ws = self.target_wb.Worksheets(self.target_sheet)
                logger.debug(f"ターゲットシートを取得: {self.target_sheet}")
            except Exception as e:
                raise GoogleSheetsError(
                    f"ターゲットシートが見つかりません: {self.target_sheet}\n\n"
                    f"ファイル: {self.target_filename}",
                    operation="Excel接続",
                    original_error=e
                ) from e

            # 接続状態の検証
            if self.excel is None or self.target_ws is None:
                raise GoogleSheetsError(
                    "Excelへの接続に失敗しました。\n"
                    "ワークシートまたはアプリケーションオブジェクトが取得できませんでした。",
                    operation="Excel接続"
                )

        except GoogleSheetsError:
            raise
        except Exception as e:
            raise GoogleSheetsError(
                f"Excelへの接続に失敗しました: {e}",
                operation="Excel接続",
                original_error=e
            ) from e

    def _cleanup_resources(self) -> None:
        """
        リソースをクリーンアップ

        Note:
            Google Sheetsリーダーと Excel COMオブジェクトを解放
        """
        logger.debug("リソースをクリーンアップ中...")

        # Google Sheetsリーダーを解放
        self.sheets_reader = None

        # Excel COMオブジェクトをクリーンアップ
        self.target_ws = None
        self.target_wb = None

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
        Google Sheets → Excel転記処理を実行

        Raises:
            GoogleSheetsError: 転記処理中にエラーが発生した場合
            SystemExit: システム終了が要求された場合
            KeyboardInterrupt: キーボード割り込みが発生した場合
        """
        logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)
        logger.info("Google Sheets → Excel自動転記処理を開始")
        logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)

        try:
            # Google Sheetsに接続
            self._report_progress("Google Sheetsに接続中...")
            self._connect_to_sheets()

            # ターゲットExcelに接続
            self._report_progress("ターゲットExcelファイルに接続中...")
            self._connect_to_target_excel()
            logger.info("ターゲットExcelファイルに接続しました")

            # 3つのループを実行（基底クラスの共通処理）
            self._execute_transfer_loops()

            logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)
            logger.info("Google Sheets → Excel自動転記処理が完了しました")
            logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)
            self._report_progress("Google Sheets転記処理が完了しました")

        except GoogleSheetsError:
            raise
        except (SystemExit, KeyboardInterrupt):
            logger.warning("処理が中断されました")
            raise
        except Exception as e:
            logger.error(f"転記処理中にエラーが発生: {e}", exc_info=True)
            raise GoogleSheetsError(
                f"転記処理中にエラーが発生しました\n"
                f"参照: {self.sheets_url} ({self.ref_sheet})\n"
                f"対象: {self.target_filename} ({self.target_sheet})\n"
                f"エラー: {e}",
                operation="転記処理",
                original_error=e
            ) from e
        finally:
            # 必ずクリーンアップを実行
            self._cleanup_resources()
