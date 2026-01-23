"""
Excel自動転記モジュール

年間行事計画（参照ファイル）から様式ファイル（反映ファイル）へ
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

from base_excel_transfer import BaseExcelTransfer
from exceptions import PDFMergeError
from constants import ExcelLookIn, ExcelLookAt, ExcelTransferConstants, PDFConversionConstants

# ロガーの設定
logger = logging.getLogger(__name__)


class ExcelTransferError(PDFMergeError):
    """Excel転記処理エラー"""
    pass


class ExcelTransfer(BaseExcelTransfer):
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
        super().__init__(target_filename, target_sheet, progress_callback, cancel_check)

        self.ref_filename = ref_filename
        self.ref_sheet = ref_sheet

        # 参照ファイル用のCOMオブジェクト
        self.ref_wb: Any = None
        self.ref_ws: Any = None

    def _get_error_class(self) -> type:
        """
        エラークラスを取得

        Returns:
            type: ExcelTransferError
        """
        return ExcelTransferError

    def _read_data_row(self, found_row: int, start_col: str, end_col: str) -> list:
        """
        参照Excelから行データを読み取る

        Args:
            found_row: 行番号
            start_col: 開始列
            end_col: 終了列

        Returns:
            list: 行データ
        """
        range_str = f"{start_col}{found_row}:{end_col}{found_row}"
        rng = self.ref_ws.Range(range_str).Value

        if rng is None:
            return []

        # Excelは2次元タプルで返すので、1次元リストに変換
        return list(rng[0]) if isinstance(rng, tuple) else [rng]

    def _find_value_in_source(self, search_value: str) -> Optional[int]:
        """
        参照ExcelのC列から値を検索

        Args:
            search_value: 検索値

        Returns:
            Optional[int]: 見つかった行番号（見つからない場合None）
        """
        ref_col_C = self.ref_ws.Columns(ExcelTransferConstants.REF_SEARCH_COL)
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
                if ref_basename in wb.Name:
                    self.ref_wb = wb
                    logger.debug(f"参照ファイルを検出: {wb.Name}")
                if target_basename in wb.Name:
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
                self.ref_ws = self.ref_wb.Worksheets(self.ref_sheet)
                logger.debug(f"参照シートを取得: {self.ref_sheet}")
            except Exception as e:
                raise ExcelTransferError(
                    f"参照シートが見つかりません: {self.ref_sheet}\n\n"
                    f"ファイル: {self.ref_filename}",
                    operation="Excel接続",
                    original_error=e
                ) from e

            try:
                self.target_ws = self.target_wb.Worksheets(self.target_sheet)
                logger.debug(f"反映シートを取得: {self.target_sheet}")
            except Exception as e:
                raise ExcelTransferError(
                    f"反映シートが見つかりません: {self.target_sheet}\n\n"
                    f"ファイル: {self.target_filename}",
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

            # 3つのループを実行（基底クラスの共通処理）
            self._execute_transfer_loops()

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


def main() -> None:
    """
    メイン処理（GUIから呼び出される）

    設定ファイルから値を取得して実行
    """
    from config_loader import ConfigLoader

    # 設定ファイルから取得
    config = ConfigLoader()
    ref_filename = config.get('files', 'excel_reference')
    target_filename = config.get('files', 'excel_target')
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
