"""
Office変換モジュール

Word、Excel、PowerPointファイルをPDFに変換する機能を提供
"""
import logging
import os
import shutil
import subprocess
from contextlib import contextmanager
from typing import Any, Optional, Generator

from win32com import client
import pythoncom
import win32process

from exceptions import PDFConversionError
from constants import WordFormat, ExcelFormat, PowerPointFormat, PDFConversionConstants

logger = logging.getLogger(__name__)


class OfficeConverter:
    """OfficeファイルをPDFに変換するクラス"""

    OFFICE_EXTENSIONS = ('.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.rtf')

    def __init__(self, temp_dir: str) -> None:
        """
        Args:
            temp_dir: 一時ファイルの保存先ディレクトリ
        """
        self.temp_dir = temp_dir

    @staticmethod
    @contextmanager
    def _com_context() -> Generator[None, None, None]:
        """COMオブジェクト用のコンテキストマネージャー"""
        pythoncom.CoInitialize()
        try:
            yield
        finally:
            pythoncom.CoUninitialize()

    @staticmethod
    def _kill_office_process(process_name: str, process_id: Optional[int] = None) -> None:
        """Officeプロセスを強制終了（可能な限りPID指定）"""
        try:
            command = ['taskkill', '/F', '/PID', str(process_id)] if process_id is not None else ['taskkill', '/F', '/IM', process_name]
            result = subprocess.run(
                command,
                capture_output=True,
                timeout=5,
                text=True
            )
            if result.returncode == 0:
                if process_id is not None:
                    logger.debug(f"プロセスを強制終了: {process_name} (pid={process_id})")
                else:
                    logger.debug(f"プロセスを強制終了: {process_name}")
            elif result.returncode == 128:
                # プロセスが見つからない（既に終了済み）
                if process_id is not None:
                    logger.debug(f"プロセスは既に終了済み: {process_name} (pid={process_id})")
                else:
                    logger.debug(f"プロセスは既に終了済み: {process_name}")
            else:
                logger.warning(
                    f"プロセス強制終了に失敗 ({process_name}): "
                    f"戻り値={result.returncode}, stderr={result.stderr}"
                )
        except subprocess.TimeoutExpired:
            logger.warning(f"プロセス強制終了がタイムアウト ({process_name}): taskkillが応答しません")
        except Exception as e:
            logger.warning(f"プロセス強制終了に失敗 ({process_name}): {e}")

    @staticmethod
    def _get_process_id(app: Any) -> Optional[int]:
        """COMアプリケーションからPIDを取得する。"""
        for hwnd_attr in ("Hwnd", "HWND"):
            try:
                hwnd = int(getattr(app, hwnd_attr))
                if hwnd:
                    _, process_id = win32process.GetWindowThreadProcessId(hwnd)
                    if process_id:
                        return process_id
            except Exception:
                continue
        return None

    def _cleanup_office_app(
        self,
        document: Any,
        app: Any,
        process_name: str,
        app_name: str,
        process_id: Optional[int] = None
    ) -> None:
        """
        Officeアプリケーションのクリーンアップを行う共通メソッド

        Args:
            document: ドキュメントオブジェクト（Word Doc, Excel Workbook, PowerPoint Presentation）
            app: アプリケーションオブジェクト
            process_name: プロセス名（強制終了用）
            app_name: アプリケーション名（ログ用）
        """
        quit_success = False

        # ドキュメントのクローズ
        if document is not None:
            try:
                # PowerPointの場合は引数なしでClose()を呼び出す
                if app_name == "PowerPoint":
                    document.Close()
                else:
                    document.Close(SaveChanges=False)
            except Exception as e:
                logger.warning(f"{app_name}ドキュメントのクローズに失敗: {e}")

        # アプリケーションの終了
        if app is not None:
            try:
                app.Quit()
                quit_success = True
            except Exception as e:
                logger.warning(f"{app_name}アプリケーションの終了に失敗: {e}")

            # Quit失敗時は、PID指定でのみ強制終了（他インスタンス巻き込みを防止）
            if not quit_success:
                if process_id is not None:
                    self._kill_office_process(process_name, process_id)
                else:
                    logger.warning(
                        f"{app_name}のPIDを取得できなかったため、"
                        "プロセス強制終了をスキップしました"
                    )

    def convert(self, file_path: str, output_path: str) -> Optional[str]:
        """
        Officeファイルを変換

        Args:
            file_path: 変換元Officeファイルのパス
            output_path: 出力先PDFのパス

        Returns:
            変換後のPDFパス（成功時）

        Raises:
            PDFConversionError: 変換処理中にエラーが発生した場合
            SystemExit, KeyboardInterrupt: システム終了・キーボード割り込み
        """
        try:
            os.remove(output_path)
        except FileNotFoundError:
            pass  # File doesn't exist, which is fine

        ext = os.path.splitext(file_path)[1].lower()
        file_name = os.path.basename(file_path)

        try:
            if ext in ['.doc', '.docx', '.rtf']:
                logger.info(f"Word変換開始: {file_name}")
                self._convert_word(file_path, output_path)
            elif ext in ['.xls', '.xlsx']:
                logger.info(f"Excel変換開始: {file_name}")
                self._convert_excel(file_path, output_path)
            elif ext in ['.ppt', '.pptx']:
                logger.info(f"PowerPoint変換開始: {file_name}")
                self._convert_powerpoint(file_path, output_path)

            if os.path.exists(output_path):
                logger.info(f"変換完了: {file_name}")
                return output_path
            else:
                logger.error(f"変換後のファイルが見つかりません: {output_path}")
                return None

        except (SystemExit, KeyboardInterrupt):
            raise

        except PDFConversionError:
            raise

        except Exception as e:
            logger.exception(f"Office変換エラー ({file_path}): {e}")
            raise PDFConversionError(f"Office変換に失敗: {file_path}", original_error=e) from e

    def _convert_word(self, file_path: str, output_path: str) -> None:
        """Word文書をPDFに変換"""
        # COMオブジェクトはバックスラッシュのパスを要求するため正規化
        file_path = os.path.normpath(file_path)
        output_path = os.path.normpath(output_path)

        with self._com_context():
            word: Any = None
            doc: Any = None
            process_id: Optional[int] = None
            try:
                word = client.DispatchEx("Word.Application")
                process_id = self._get_process_id(word)
                try:
                    word.Visible = False
                except Exception as e:
                    logger.debug(f"Word.Visible設定をスキップ: {e}")
                word.DisplayAlerts = False
                doc = word.Documents.Open(file_path, ReadOnly=True)
                doc.SaveAs2(output_path, FileFormat=WordFormat.PDF)
                logger.debug(f"Word変換完了: {file_path} -> {output_path}")
            finally:
                self._cleanup_office_app(
                    doc, word, "WINWORD.EXE", "Word", process_id=process_id
                )

    def _convert_excel(self, file_path: str, output_path: str) -> None:
        """ExcelワークブックをPDFに変換"""
        # COMオブジェクトはバックスラッシュのパスを要求するため正規化
        file_path = os.path.normpath(file_path)
        output_path = os.path.normpath(output_path)

        # ネットワークパスの場合はローカルにコピー
        base_name = os.path.basename(file_path)
        local_copy = os.path.join(self.temp_dir, PDFConversionConstants.LOCAL_COPY_PREFIX + base_name)

        try:
            shutil.copy2(file_path, local_copy)
        except OSError as e:
            logger.error(f"ファイルコピーに失敗 ({file_path}): {e}")
            raise PDFConversionError(f"ファイルコピーに失敗: {file_path}", original_error=e) from e

        with self._com_context():
            excel: Any = None
            wb: Any = None
            process_id: Optional[int] = None
            try:
                excel = client.DispatchEx("Excel.Application")
                process_id = self._get_process_id(excel)
                try:
                    excel.Visible = False
                except Exception as e:
                    logger.debug(f"Excel.Visible設定をスキップ: {e}")
                excel.DisplayAlerts = False
                wb = excel.Workbooks.Open(local_copy, ReadOnly=True)
                wb.ExportAsFixedFormat(ExcelFormat.PDF, output_path)
                logger.debug(f"Excel変換完了: {file_path} -> {output_path}")
            finally:
                self._cleanup_office_app(
                    wb, excel, "EXCEL.EXE", "Excel", process_id=process_id
                )
                # ローカルコピーを削除
                if os.path.exists(local_copy):
                    try:
                        os.remove(local_copy)
                    except OSError as e:
                        logger.warning(f"ローカルコピーの削除に失敗 ({local_copy}): {e}")

    def _convert_powerpoint(self, file_path: str, output_path: str) -> None:
        """PowerPointプレゼンテーションをPDFに変換"""
        # COMオブジェクトはバックスラッシュのパスを要求するため正規化
        file_path = os.path.normpath(file_path)
        output_path = os.path.normpath(output_path)

        with self._com_context():
            powerpoint: Any = None
            pres: Any = None
            process_id: Optional[int] = None
            try:
                powerpoint = client.DispatchEx("PowerPoint.Application")
                process_id = self._get_process_id(powerpoint)
                # PowerPointはVisible=Falseをサポートしない環境があるため、
                # WithWindow=Falseのみ使用してウィンドウを非表示化
                logger.debug("PowerPointを起動 (WithWindow=False)")
                pres = powerpoint.Presentations.Open(file_path, WithWindow=False)
                pres.SaveAs(output_path, PowerPointFormat.PDF)
                logger.debug(f"PowerPoint変換完了: {file_path} -> {output_path}")
            finally:
                self._cleanup_office_app(
                    pres, powerpoint, "POWERPNT.EXE", "PowerPoint", process_id=process_id
                )
