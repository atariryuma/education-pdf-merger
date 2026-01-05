"""
PDF変換モジュール

各種ファイル（Office、画像、一太郎、PDF）をPDFに変換する機能を提供
"""
import hashlib
import logging
import os
import re
import time
import shutil
import subprocess
from contextlib import contextmanager
from typing import Dict, List, Optional, Any, Generator, Callable, Tuple

from win32com import client
import pythoncom
from PIL import Image
from pywinauto import Application
from pywinauto.keyboard import send_keys

# 例外クラスと定数は専用モジュールからインポート
from exceptions import PDFConversionError, CancelledError
from constants import MSOfficeConstants, AppConstants, PDFConversionConstants
from path_validator import PathValidator

# ロガーの設定
logger = logging.getLogger(__name__)


class PDFConverter:
    """各種ファイルをPDFに変換するクラス"""

    # サポートされる拡張子
    OFFICE_EXTENSIONS: List[str] = ['.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.rtf']
    IMAGE_EXTENSIONS: List[str] = ['.jpg', '.jpeg', '.png', '.bmp', '.tiff']
    ICHITARO_EXTENSIONS: List[str] = ['.jtd']
    PDF_EXTENSION: str = '.pdf'

    # Office FileFormat定数（constants.pyから参照）
    WORD_PDF_FORMAT: int = MSOfficeConstants.WORD_PDF_FORMAT
    EXCEL_PDF_FORMAT: int = MSOfficeConstants.EXCEL_PDF_FORMAT
    POWERPOINT_PDF_FORMAT: int = MSOfficeConstants.POWERPOINT_PDF_FORMAT

    # 一太郎変換のデフォルトタイミング設定（constants.pyから参照）
    DEFAULT_ICHITARO_SETTINGS: Dict[str, Any] = {
        **AppConstants.ICHITARO_DEFAULTS,
        'printer_name': 'Microsoft Print to PDF'
    }

    def __init__(
        self,
        temp_dir: str,
        ichitaro_settings: Optional[Dict[str, Any]] = None,
        cancel_check: Optional[Callable[[], bool]] = None,
        dialog_callback: Optional[Callable[[str, bool], None]] = None,
        config: Optional[Any] = None
    ) -> None:
        """
        Args:
            temp_dir: 一時ファイルの保存先ディレクトリ
            ichitaro_settings: 一太郎変換のタイミング設定（オプション）
            cancel_check: キャンセル状態をチェックするコールバック関数
            dialog_callback: 一太郎変換ダイアログのコールバック関数(message, show)
            config: ConfigLoaderインスタンス（区切りページ生成に必要）
        """
        self.temp_dir = temp_dir
        # 一太郎設定をマージ（指定されたものでデフォルトを上書き）
        self.ichitaro_settings: Dict[str, Any] = self.DEFAULT_ICHITARO_SETTINGS.copy()
        if ichitaro_settings:
            self.ichitaro_settings.update(ichitaro_settings)
        self._cancel_check = cancel_check or (lambda: False)
        self._dialog_callback = dialog_callback
        self.config = config

    def is_cancelled(self) -> bool:
        """キャンセルされたかどうかを確認"""
        return self._cancel_check()

    def _wait_with_cancel_check(self, seconds: float) -> None:
        """
        キャンセルチェック付きの待機

        Args:
            seconds: 待機時間（秒）

        Raises:
            CancelledError: キャンセルされた場合
        """
        interval = PDFConversionConstants.CANCEL_CHECK_INTERVAL
        elapsed = 0.0
        while elapsed < seconds:
            if self.is_cancelled():
                logger.info("一太郎変換がキャンセルされました")
                self._cleanup_ichitaro_windows()
                raise CancelledError("一太郎変換がキャンセルされました")
            wait_time = min(interval, seconds - elapsed)
            time.sleep(wait_time)
            elapsed += wait_time

    @staticmethod
    def _generate_unique_filename(file_path: str, extension: str = '.pdf') -> str:
        """
        ファイルパスからユニークなファイル名を生成

        Args:
            file_path: 元のファイルパス
            extension: 出力ファイルの拡張子

        Returns:
            str: ユニークなファイル名（パスのハッシュを含む）
        """
        base_name = os.path.basename(file_path)
        name_without_ext = os.path.splitext(base_name)[0]
        # パス全体のハッシュを使って衝突を回避（SHA256を使用）
        path_hash = hashlib.sha256(file_path.encode('utf-8')).hexdigest()[:8]
        return f"{name_without_ext}_{path_hash}{extension}"

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
    def _is_temporary_file(file_path: str) -> bool:
        """一時ファイルかどうかを判定

        - ~$ を含むファイル名（Office一時ファイル）
        - .$ で始まる拡張子（一太郎一時ファイル: .$td など）
        """
        base_name = os.path.basename(file_path)
        ext = os.path.splitext(file_path)[1].lower()
        # ~$を含むファイルまたは.$で始まる拡張子
        return '~$' in base_name or ext.startswith('.$')

    @staticmethod
    def _escape_for_send_keys(text: str) -> str:
        """
        send_keys用に特殊文字をエスケープ

        Args:
            text: エスケープする文字列

        Returns:
            str: エスケープされた文字列
        """
        # pywinautoのsend_keysで特殊な意味を持つ文字をエスケープ
        # 注意: { と } は最初に処理し、他の文字より先にエスケープする必要がある
        # 一文字ずつ処理して正しくエスケープする
        special_chars = {'{', '}', '+', '^', '%', '~', '(', ')'}
        result = []
        for char in text:
            if char in special_chars:
                result.append('{')
                result.append(char)
                result.append('}')
            else:
                result.append(char)
        return ''.join(result)

    def convert(self, file_path: str, output_path: Optional[str] = None) -> Optional[str]:
        """
        ファイルをPDFに変換

        Args:
            file_path: 変換元ファイルのパス
            output_path: 出力先PDFのパス（省略時は自動生成）

        Returns:
            str: 変換後のPDFパス（失敗時はNone）

        Raises:
            PDFConversionError: 変換処理中にエラーが発生した場合
        """
        # 一時ファイルをスキップ
        if self._is_temporary_file(file_path):
            logger.info(f"一時ファイルをスキップ: {os.path.basename(file_path)}")
            return None

        ext = os.path.splitext(file_path)[1].lower()

        # 出力パスの決定（衝突回避のためハッシュを含む）
        if output_path is None:
            unique_name = self._generate_unique_filename(file_path)
            output_path = os.path.join(self.temp_dir, unique_name)

        # 既に変換済みの場合はスキップ
        if os.path.exists(output_path):
            logger.info(f"変換済みファイルが存在: {output_path}")
            return output_path

        # 拡張子に応じて変換
        if ext in self.OFFICE_EXTENSIONS:
            logger.info(f"Officeファイルを変換: {file_path}")
            return self._convert_office(file_path, output_path)
        elif ext == self.PDF_EXTENSION:
            logger.debug(f"PDFファイル: {file_path}")
            return file_path if os.path.exists(file_path) else None
        elif ext in self.IMAGE_EXTENSIONS:
            logger.info(f"画像ファイルを変換: {file_path}")
            return self._convert_image(file_path, output_path)
        elif ext in self.ICHITARO_EXTENSIONS:
            logger.info(f"一太郎ファイルを変換: {file_path}")
            return self._convert_ichitaro(file_path, output_path)
        else:
            logger.warning(f"サポートされていないファイル形式: {file_path}")
            return None

    def _convert_office(self, file_path: str, output_path: str) -> Optional[str]:
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
        if os.path.exists(output_path):
            os.remove(output_path)

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
            # システム終了・キーボード割り込みは再スロー
            raise

        except PDFConversionError:
            # 既にPDFConversionErrorの場合は再スロー
            raise

        except Exception as e:
            logger.exception(f"Office変換エラー ({file_path}): {e}")
            raise PDFConversionError(f"Office変換に失敗: {file_path}") from e

    @staticmethod
    def _kill_office_process(process_name: str) -> None:
        """Officeプロセスを強制終了（フォールバック）"""
        try:
            subprocess.run(
                ['taskkill', '/F', '/IM', process_name],
                capture_output=True,
                timeout=5
            )
            logger.debug(f"プロセスを強制終了: {process_name}")
        except Exception as e:
            logger.warning(f"プロセス強制終了に失敗 ({process_name}): {e}")

    def _cleanup_office_app(
        self,
        document: Any,
        app: Any,
        process_name: str,
        app_name: str
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

            # Quit失敗時はプロセスを強制終了
            if not quit_success:
                self._kill_office_process(process_name)

    def _convert_word(self, file_path: str, output_path: str) -> None:
        """Word文書をPDFに変換"""
        # COMオブジェクトはバックスラッシュのパスを要求するため正規化
        file_path = os.path.normpath(file_path)
        output_path = os.path.normpath(output_path)

        with self._com_context():
            word: Any = None
            doc: Any = None
            try:
                word = client.Dispatch("Word.Application")
                try:
                    word.Visible = False
                except Exception as e:
                    logger.debug(f"Word.Visible設定をスキップ: {e}")
                word.DisplayAlerts = False
                doc = word.Documents.Open(file_path, ReadOnly=True)
                doc.SaveAs2(output_path, FileFormat=self.WORD_PDF_FORMAT)
                logger.debug(f"Word変換完了: {file_path} -> {output_path}")
            finally:
                self._cleanup_office_app(doc, word, "WINWORD.EXE", "Word")

    def _convert_excel(self, file_path: str, output_path: str) -> None:
        """ExcelワークブックをPDFに変換"""
        # COMオブジェクトはバックスラッシュのパスを要求するため正規化
        file_path = os.path.normpath(file_path)
        output_path = os.path.normpath(output_path)

        # ネットワークパスの場合はローカルにコピー
        base_name = os.path.basename(file_path)
        local_copy = os.path.join(self.temp_dir, "local_copy_" + base_name)

        try:
            shutil.copy2(file_path, local_copy)
        except OSError as e:
            logger.error(f"ファイルコピーに失敗 ({file_path}): {e}")
            raise PDFConversionError(f"ファイルコピーに失敗: {file_path}") from e

        with self._com_context():
            excel: Any = None
            wb: Any = None
            try:
                excel = client.Dispatch("Excel.Application")
                try:
                    excel.Visible = False
                except Exception as e:
                    logger.debug(f"Excel.Visible設定をスキップ: {e}")
                excel.DisplayAlerts = False
                wb = excel.Workbooks.Open(local_copy, ReadOnly=True)
                wb.ExportAsFixedFormat(self.EXCEL_PDF_FORMAT, output_path)
                logger.debug(f"Excel変換完了: {file_path} -> {output_path}")
            finally:
                self._cleanup_office_app(wb, excel, "EXCEL.EXE", "Excel")
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
            try:
                powerpoint = client.Dispatch("PowerPoint.Application")
                # PowerPointはVisible=Falseをサポートしない環境があるため、
                # WithWindow=Falseのみ使用してウィンドウを非表示化
                logger.debug("PowerPointを起動 (WithWindow=False)")
                pres = powerpoint.Presentations.Open(file_path, WithWindow=False)
                pres.SaveAs(output_path, self.POWERPOINT_PDF_FORMAT)
                logger.debug(f"PowerPoint変換完了: {file_path} -> {output_path}")
            finally:
                self._cleanup_office_app(pres, powerpoint, "POWERPNT.EXE", "PowerPoint")

    def _convert_image(self, file_path: str, output_path: str) -> Optional[str]:
        """
        画像ファイルをPDFに変換

        Args:
            file_path: 変換元画像ファイルのパス
            output_path: 出力先PDFのパス

        Returns:
            変換後のPDFパス（成功時）

        Raises:
            PDFConversionError: 変換処理中にエラーが発生した場合
        """
        try:
            # with文でリソースを確実に解放
            with Image.open(file_path) as image:
                # RGBA/Pモードの場合はRGBに変換（新しいオブジェクトが返る）
                if image.mode in ("RGBA", "P"):
                    image = image.convert("RGB")
                image.save(output_path, "PDF")

            logger.debug(f"画像変換完了: {file_path} -> {output_path}")
            return output_path
        except IOError as e:
            logger.error(f"画像変換エラー ({file_path}): {e}")
            raise PDFConversionError(f"画像変換に失敗: {file_path}") from e

    def _convert_ichitaro(self, file_path: str, output_path: str) -> Optional[str]:
        """
        一太郎ファイルをPDFに変換（最大3回試行）

        処理フロー:
        1. 既存の一太郎プロセスをクリーンアップ
        2. 一太郎でファイルを開く
        3. 印刷→保存操作
        4. PDF作成完了を待つ
        5. 一太郎を正常終了
        ※ 失敗時は最大3回まで試行（初回を含む）

        Args:
            file_path: 一太郎ファイルのパス
            output_path: 出力先PDFのパス

        Returns:
            変換後のPDFパス（成功時）、失敗時はNone

        Raises:
            CancelledError: ユーザーがキャンセルした場合
        """
        max_attempts = PDFConversionConstants.ICHITARO_MAX_ATTEMPTS
        file_name = os.path.basename(file_path)

        # ダイアログ表示（ファイル変換の最初に1回だけ）
        if self._dialog_callback:
            self._dialog_callback(f"変換中: {file_name}", True)  # True = show

        try:
            for attempt in range(1, max_attempts + 1):
                logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)
                if attempt > 1:
                    logger.info(f"一太郎PDF変換を再試行します（試行 {attempt}/{max_attempts}）")
                logger.info(f"一太郎PDF変換開始: {file_name}")
                logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)

                # パスを正規化（バックスラッシュに統一）
                file_path_norm = os.path.normpath(file_path)
                output_path_norm = os.path.normpath(output_path)

                logger.info(f"入力ファイル: {file_path_norm}")
                logger.info(f"出力ファイル: {output_path_norm}")

                # 既存の出力ファイルを削除
                if os.path.exists(output_path_norm):
                    logger.info(f"既存の出力ファイルを削除: {output_path_norm}")
                    os.remove(output_path_norm)

                # 設定値
                max_wait = self.ichitaro_settings.get('ichitaro_ready_timeout', 30)
                save_wait = self.ichitaro_settings.get('save_wait_seconds', 20)

                logger.info(f"設定値 - 接続タイムアウト: {max_wait}秒, 保存待機: {save_wait}秒")

                app = None
                main_window = None

                try:
                    # ステップ1: 事前クリーンアップ
                    logger.info(PDFConversionConstants.LOG_SEPARATOR_MINOR)
                    logger.info("ステップ1: 事前クリーンアップ")
                    logger.info(PDFConversionConstants.LOG_SEPARATOR_MINOR)
                    self._cleanup_ichitaro_windows()

                    # ステップ2: 一太郎でファイルを開く
                    logger.info(PDFConversionConstants.LOG_SEPARATOR_MINOR)
                    logger.info("ステップ2: 一太郎でファイルを開く")
                    logger.info(PDFConversionConstants.LOG_SEPARATOR_MINOR)
                    app, main_window = self._open_ichitaro_file(file_path_norm, max_wait)
                    if app is None or main_window is None:
                        logger.error("一太郎ファイルを開けませんでした")
                        result = None
                    else:
                        # ステップ3: 印刷→保存操作
                        logger.info(PDFConversionConstants.LOG_SEPARATOR_MINOR)
                        logger.info("ステップ3: 印刷→保存操作")
                        logger.info(PDFConversionConstants.LOG_SEPARATOR_MINOR)
                        self._execute_print_sequence(app, output_path_norm)

                        # ステップ4: PDF作成完了を待つ
                        logger.info(PDFConversionConstants.LOG_SEPARATOR_MINOR)
                        logger.info("ステップ4: PDF作成完了を待つ")
                        logger.info(PDFConversionConstants.LOG_SEPARATOR_MINOR)
                        result = self._wait_for_output_file(output_path_norm, file_path_norm, save_wait)

                finally:
                    # ステップ5: 一太郎を正常終了
                    logger.info(PDFConversionConstants.LOG_SEPARATOR_MINOR)
                    logger.info("ステップ5: 一太郎を正常終了")
                    logger.info(PDFConversionConstants.LOG_SEPARATOR_MINOR)
                    self._close_ichitaro(app, main_window)

                # 最終確認
                if result and os.path.exists(output_path_norm):
                    file_size = os.path.getsize(output_path_norm)
                    if file_size > 0:
                        logger.info(f"{PDFConversionConstants.LOG_MARK_SUCCESS} PDFファイル確認OK (サイズ: {file_size:,} bytes)")
                        logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)
                        if attempt > 1:
                            logger.info(f"{PDFConversionConstants.LOG_MARK_SUCCESS} 一太郎PDF変換成功（試行 {attempt}/{max_attempts}）: {file_name}")
                        else:
                            logger.info(f"{PDFConversionConstants.LOG_MARK_SUCCESS} 一太郎PDF変換成功: {file_name}")
                        logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)
                        return result  # 成功したら即座にreturn
                    else:
                        logger.error(f"{PDFConversionConstants.LOG_MARK_FAILURE} PDFファイルが空です (サイズ: 0 bytes)")
                        result = None
                elif result:
                    logger.error(f"{PDFConversionConstants.LOG_MARK_FAILURE} PDFファイルが削除されています: {output_path_norm}")
                    result = None

                # 変換失敗時の処理
                logger.error(PDFConversionConstants.LOG_SEPARATOR_MAJOR)
                logger.error(f"{PDFConversionConstants.LOG_MARK_FAILURE} 一太郎PDF変換失敗（試行 {attempt}/{max_attempts}）: {file_name}")
                logger.error(PDFConversionConstants.LOG_SEPARATOR_MAJOR)

                if attempt < max_attempts:
                    logger.warning("一太郎プロセスをクリーンアップして再試行します...")
                    self._cleanup_ichitaro_windows()
                    self._wait_with_cancel_check(PDFConversionConstants.ICHITARO_RETRY_DELAY)
                    # ダイアログメッセージ更新
                    if self._dialog_callback:
                        self._dialog_callback(f"再試行中: {file_name} ({attempt+1}/{max_attempts})", True)
                else:
                    logger.error(f"{PDFConversionConstants.LOG_MARK_FAILURE} 一太郎変換が{max_attempts}回失敗しました。スキップします: {file_name}")

            # ループを抜けた = 全試行失敗
            logger.warning(f"一太郎変換が全{max_attempts}回の試行で失敗しました: {file_name}")
            return None

        except CancelledError:
            # キャンセルされた場合は即座に終了
            logger.info("一太郎変換がキャンセルされました")
            raise

        except (SystemExit, KeyboardInterrupt):
            # システム終了・キーボード割り込みは再スロー
            raise

        except Exception as e:
            # 予期しないエラーはログに記録してNoneを返す
            logger.exception(f"{PDFConversionConstants.LOG_MARK_FAILURE} 一太郎変換で予期しないエラー: {e}")
            return None

        finally:
            # ダイアログ非表示（ファイル変換の最後に1回だけ）
            if self._dialog_callback:
                self._dialog_callback("", False)  # False = hide

    def _open_ichitaro_file(
        self,
        file_path: str,
        max_wait: int
    ) -> Tuple[Optional[Application], Optional[Any]]:
        """
        一太郎でファイルを開く（ベストプラクティス版）

        Args:
            file_path: 一太郎ファイルのパス
            max_wait: 最大待機時間（秒）

        Returns:
            (app, main_window): ApplicationオブジェクトとメインウィンドウのTuple（成功時）
            (None, None): 接続失敗時

        Raises:
            CancelledError: キャンセルされた場合
        """
        file_name = os.path.basename(file_path)

        if self.is_cancelled():
            raise CancelledError("一太郎変換がキャンセルされました")

        try:
            # ファイルを開く
            logger.info(f"一太郎でファイルを開く: {file_name}")
            os.startfile(file_path)

            # 一太郎起動待機（キャンセル可能）
            wait_time = PDFConversionConstants.ICHITARO_STARTUP_WAIT
            logger.info(f"一太郎の起動を待機中（{wait_time}秒）...")
            self._wait_with_cancel_check(wait_time)

            # ファイル名ベースのウィンドウ検索
            escaped_name = re.escape(file_name)
            title_pattern = f".*{escaped_name}.*"
            logger.info(f"一太郎ウィンドウに接続中: '{title_pattern}'")

            # pywinautoで接続（タイムアウト付き）
            app = Application(backend="uia").connect(
                title_re=title_pattern, timeout=max_wait
            )
            main_window = app.top_window()
            logger.info(f"{PDFConversionConstants.LOG_MARK_SUCCESS} 一太郎への接続成功: {main_window.window_text()}")
            main_window.set_focus()
            return app, main_window

        except CancelledError:
            self._cleanup_ichitaro_windows()
            raise
        except Exception as e:
            logger.error(f"{PDFConversionConstants.LOG_MARK_FAILURE} 一太郎への接続失敗: {e}")
            self._cleanup_ichitaro_windows()
            return None, None

    def _close_ichitaro(
        self,
        app: Optional[Application],
        main_window: Optional[Any]
    ) -> None:
        """
        一太郎を正常終了（改良版）

        ファイルが変更されている場合でも保存確認なしで閉じる：
        1. 印刷処理完了を待つ
        2. app.kill()で強制終了（保存確認ダイアログを回避）

        Args:
            app: pywinauto Applicationオブジェクト
            main_window: メインウィンドウオブジェクト
        """
        if app is None or main_window is None:
            logger.info("一太郎が開いていないため、クローズ不要")
            return

        try:
            # 印刷処理完了を待つ
            wait_time = PDFConversionConstants.ICHITARO_PRINT_COMPLETE_WAIT
            logger.info(f"印刷処理の完全終了を待機中（{wait_time}秒）...")
            self._wait_with_cancel_check(wait_time)

            # 一太郎を強制終了（保存確認ダイアログを回避）
            # window.close()だとファイル変更時に保存確認が出るため、app.kill()を使用
            logger.info("一太郎プロセスを終了中（app.kill()）...")
            window_title = main_window.window_text()
            logger.info(f"対象ウィンドウ: {window_title}")

            app.kill()
            logger.info(f"{PDFConversionConstants.LOG_MARK_SUCCESS} 一太郎プロセスを終了しました")

            # プロセス終了の確認待機
            wait_time = PDFConversionConstants.ICHITARO_WINDOW_CLOSE_WAIT
            self._wait_with_cancel_check(wait_time)

        except Exception as e:
            logger.warning(f"一太郎の終了に失敗しました: {e}")
            # 既に終了している可能性があるため、エラーでも続行

    def _execute_print_sequence(
        self,
        app: Application,
        output_path: str
    ) -> bool:
        """
        印刷シーケンスを実行（ベストプラクティス版）

        手順：
        Ctrl+P → Microsoft Print to PDFを選択 → Enter → ファイル名入力 → Enter

        Args:
            app: pywinauto Applicationオブジェクト
            output_path: 出力PDFファイルのパス

        Returns:
            bool: 処理成功時True（常にTrue、失敗時は例外）

        Raises:
            CancelledError: キャンセルされた場合
            Exception: プリンター選択に失敗した場合
        """
        # Ctrl+P で印刷ダイアログを開く（低スペックPC対応で待機時間を延長）
        logger.info("印刷ダイアログを開く (Ctrl+P)")
        send_keys("^p")
        logger.info("Ctrl+P送信完了、印刷ダイアログの表示を待機中...")
        wait_time = PDFConversionConstants.ICHITARO_CTRL_P_WAIT
        self._wait_with_cancel_check(wait_time)
        logger.info(f"{PDFConversionConstants.LOG_MARK_SUCCESS} 待機完了、印刷ダイアログが開いているはず")

        # Microsoft Print to PDFをプリンター名で直接選択（環境非依存）
        logger.info("Microsoft Print to PDFを選択中...")

        # リトライ機構：低スペックPCでダイアログの準備が遅い場合に対応
        max_retries = PDFConversionConstants.PRINTER_SELECT_MAX_RETRIES
        retry_delay = PDFConversionConstants.PRINTER_SELECT_RETRY_DELAY

        for attempt in range(max_retries):
            try:
                main_window = app.top_window()
                print_dialog = main_window.child_window(title="印刷", control_type="Window")
                printer_combo = print_dialog.child_window(auto_id="1297", control_type="ComboBox")

                # pywinautoの高レベルAPIでプリンターを選択（ユーザー操作をシミュレート）
                printer_combo.select("Microsoft Print to PDF")
                logger.info(f"{PDFConversionConstants.LOG_MARK_SUCCESS} select()メソッドで'Microsoft Print to PDF'を選択")
                self._wait_with_cancel_check(PDFConversionConstants.ICHITARO_PRINTER_SELECT_WAIT)

                # 印刷ボタンにフォーカスを設定
                try:
                    ok_button = print_dialog.child_window(title="OK", control_type="Button")
                    ok_button.set_focus()
                    logger.info("印刷ボタン（OK）にフォーカスを設定")
                    self._wait_with_cancel_check(PDFConversionConstants.ICHITARO_CTRL_A_WAIT)
                except Exception as focus_error:
                    logger.debug(f"印刷ボタンへのフォーカス設定をスキップ: {focus_error}")

                # 成功したのでリトライループを抜ける
                break

            except Exception as select_error:
                if attempt < max_retries - 1:
                    logger.warning(f"プリンター選択失敗（試行 {attempt + 1}/{max_retries}）: {select_error}")
                    logger.info(f"{retry_delay}秒待機後、再試行します...")
                    self._wait_with_cancel_check(retry_delay)
                else:
                    # 最終試行でも失敗した場合はエラーとして扱う
                    logger.error(f"プリンター選択が{max_retries}回失敗しました: {select_error}")
                    raise Exception(f"Microsoft Print to PDFの選択に失敗しました: {select_error}")

        if self.is_cancelled():
            raise CancelledError("一太郎変換がキャンセルされました")

        # Enter で印刷実行（2回押す）
        logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)
        logger.info("印刷ダイアログでEnterキーを2回押します")
        logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)
        logger.info("1回目のEnter...")
        send_keys("{ENTER}")
        self._wait_with_cancel_check(PDFConversionConstants.ICHITARO_ENTER_INTERVAL)
        logger.info("2回目のEnter...")
        send_keys("{ENTER}")
        logger.info(f"{PDFConversionConstants.LOG_MARK_SUCCESS} Enterキー2回送信完了")

        # 保存ダイアログの表示を動的に待機（複数検出方法で対応）
        logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)
        logger.info(f"保存ダイアログの表示を待機中（動的、最大{PDFConversionConstants.ICHITARO_DIALOG_TIMEOUT}秒）...")
        logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)

        # 動的ポーリング
        dialog_timeout = PDFConversionConstants.ICHITARO_DIALOG_TIMEOUT
        dialog_wait_interval = PDFConversionConstants.ICHITARO_DIALOG_POLL_INTERVAL
        dialog_elapsed = 0.0
        dialog_found = False
        min_wait = PDFConversionConstants.ICHITARO_DIALOG_MIN_WAIT

        while dialog_elapsed < dialog_timeout:
            self._wait_with_cancel_check(dialog_wait_interval)
            dialog_elapsed += dialog_wait_interval

            # 最低待機時間経過後、ダイアログ検出を試行
            if dialog_elapsed >= min_wait:
                try:
                    # 方法1: top_window()で最前面ウィンドウを取得
                    try:
                        top = app.top_window()
                        if top and top.exists(timeout=0):
                            class_name = top.class_name()

                            # 一般的な保存ダイアログのクラス名、または一太郎のメインウィンドウ以外
                            if class_name == "#32770" or "JSTARO" not in class_name.upper():
                                logger.info(f"{PDFConversionConstants.LOG_MARK_SUCCESS} 保存ダイアログ検出（top_window, {dialog_elapsed:.1f}秒経過）")
                                logger.info(f"  クラス名: {class_name}")
                                dialog_found = True
                    except Exception as e:
                        logger.debug(f"top_window検出エラー: {e}")

                    # 方法2: タイトル正規表現での検出（従来の方法）
                    if not dialog_found:
                        save_dialogs = app.windows(title_re='.*名前を付けて保存.*|.*Save.*|.*保存.*')
                        if save_dialogs:
                            for dlg in save_dialogs:
                                if dlg.exists(timeout=0):
                                    logger.info(f"{PDFConversionConstants.LOG_MARK_SUCCESS} 保存ダイアログ検出（title_re, {dialog_elapsed:.1f}秒経過）")
                                    dialog_found = True
                                    break

                    # 方法3: クラス名#32770での検出
                    if not dialog_found:
                        try:
                            dialog_windows = app.windows(class_name="#32770")
                            if dialog_windows:
                                for dlg in dialog_windows:
                                    if dlg.exists(timeout=0):
                                        logger.info(f"{PDFConversionConstants.LOG_MARK_SUCCESS} 保存ダイアログ検出（#32770, {dialog_elapsed:.1f}秒経過）")
                                        logger.info(f"  タイトル: {dlg.window_text()}")
                                        dialog_found = True
                                        break
                        except Exception as e:
                            logger.debug(f"#32770検出試行エラー: {e}")

                    if dialog_found:
                        break

                except Exception as e:
                    # ダイアログ検出失敗は無視（キーボード操作でフォールバック）
                    logger.debug(f"ダイアログ検出試行エラー: {e}")

            # 進行状況を表示（2秒ごと）
            # 2.0秒経過以降、2秒の倍数（4.0, 6.0, 8.0...）の時だけ表示
            if dialog_elapsed >= 4.0 and dialog_elapsed % 2.0 < dialog_wait_interval:
                logger.info(f"待機中... {dialog_elapsed:.1f}秒 / {dialog_timeout}秒")

        if dialog_found:
            logger.info(f"{PDFConversionConstants.LOG_MARK_SUCCESS} 保存ダイアログ確認済み")
        else:
            # ダイアログ検出できなくても続行（キーボード操作でフォールバック）
            logger.info(f"{PDFConversionConstants.LOG_MARK_SUCCESS} {dialog_elapsed:.1f}秒待機完了（ダイアログ検出なし、キーボード操作で続行）")

        # キーボード入力の準備のため追加待機
        self._wait_with_cancel_check(PDFConversionConstants.ICHITARO_KEYBOARD_PREP_WAIT)

        # ダイアログ検出をスキップして、直接キーボード操作でファイル名入力
        logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)
        logger.info("保存ダイアログへキーボード操作で直接入力")
        logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)

        logger.info("Ctrl+Aで全選択...")
        send_keys("^a")
        self._wait_with_cancel_check(PDFConversionConstants.ICHITARO_CTRL_A_WAIT)

        logger.info(f"ファイルパスを入力: {output_path}")
        escaped_path = self._escape_for_send_keys(output_path)
        logger.info(f"エスケープ済みパス: {escaped_path}")
        send_keys(escaped_path, pause=0.02, with_spaces=True)
        self._wait_with_cancel_check(PDFConversionConstants.ICHITARO_FILE_INPUT_WAIT)

        logger.info("Enterキーで保存実行...")
        send_keys("{ENTER}")
        logger.info(f"{PDFConversionConstants.LOG_MARK_SUCCESS} 保存処理完了")
        return True

    def _wait_for_output_file(
        self,
        output_path: str,
        file_path: str,
        save_wait: int = 20
    ) -> Optional[str]:
        """
        出力ファイルの作成を待機（動的間隔＆ファイルサイズ安定性チェック）

        Args:
            output_path: 出力PDFのパス
            file_path: 元ファイルのパス（ログ用）
            save_wait: 最大待機時間（秒）

        Returns:
            出力ファイルパス（成功時）、失敗時はNone

        Raises:
            CancelledError: キャンセルされた場合
        """
        logger.info(f"出力ファイルの作成を待機中（最大{save_wait}秒、動的間隔でチェック）...")
        logger.info(f"待機対象ファイル: {output_path}")

        # 動的待機間隔を生成（ジェネレータ式）
        def generate_intervals():
            """待機間隔を生成するジェネレータ"""
            for _ in range(PDFConversionConstants.FILE_WAIT_FAST_COUNT):
                yield PDFConversionConstants.FILE_WAIT_INTERVAL_FAST
            for _ in range(PDFConversionConstants.FILE_WAIT_MEDIUM_COUNT):
                yield PDFConversionConstants.FILE_WAIT_INTERVAL_MEDIUM
            # 残りは無限に低速チェック
            while True:
                yield PDFConversionConstants.FILE_WAIT_INTERVAL_SLOW

        elapsed_time = 0.0
        last_size = 0
        stable_count = 0
        last_log_time = 0.0

        for interval in generate_intervals():
            if elapsed_time > save_wait:
                break

            if os.path.exists(output_path):
                current_size = os.path.getsize(output_path)

                # ファイルサイズが0より大きく、安定している（N回連続で同じサイズ）
                if current_size > 0:
                    if current_size == last_size:
                        stable_count += 1
                        threshold = PDFConversionConstants.FILE_STABILITY_THRESHOLD
                        if stable_count >= threshold:
                            logger.info(f"{PDFConversionConstants.LOG_MARK_SUCCESS} 出力ファイル検出成功！ (サイズ: {current_size:,} bytes)")
                            logger.info(f"待機時間: {elapsed_time:.1f}秒")
                            return output_path
                    else:
                        stable_count = 0
                    last_size = current_size

            # 一定間隔ごとに経過時間をログ出力
            log_interval = PDFConversionConstants.FILE_WAIT_LOG_INTERVAL
            if elapsed_time - last_log_time >= log_interval:
                logger.info(f"待機中... 経過時間: {elapsed_time:.1f}秒 / {save_wait}秒")
                last_log_time = elapsed_time

            self._wait_with_cancel_check(interval)
            elapsed_time += interval

        logger.error(f"{PDFConversionConstants.LOG_MARK_FAILURE} タイムアウト: {save_wait}秒経過しても出力ファイルが見つかりません")
        logger.error(f"ファイルパス: {output_path}")
        return None

    def _cleanup_ichitaro_windows(self) -> None:
        """
        残っている一太郎ウィンドウをクリーンアップ（最適化版）

        Raises:
            CancelledError: キャンセルされた場合
        """
        try:
            logger.info("一太郎ウィンドウのクリーンアップを開始...")
            timeout = PDFConversionConstants.ICHITARO_CLEANUP_TIMEOUT
            app = Application(backend="uia").connect(title_re=".*一太郎.*", timeout=timeout)

            # pywinauto推奨: app.kill()で全プロセス終了
            logger.info("一太郎プロセスを強制終了しています...")
            app.kill()
            self._wait_with_cancel_check(PDFConversionConstants.ICHITARO_CLEANUP_WAIT)
            logger.info(f"{PDFConversionConstants.LOG_MARK_SUCCESS} 一太郎プロセスのクリーンアップ完了")

        except Exception:
            # 一太郎が起動していない場合は正常
            logger.info("一太郎プロセスなし（クリーンアップ不要）")

    def create_separator_page(self, folder_name: str) -> Optional[str]:
        """
        区切りページを作成（reportlab完全生成版）

        Args:
            folder_name: セクションタイトル

        Returns:
            str: 作成したPDFのパス（失敗時はNone）
        """
        try:
            # ConfigLoaderが設定されていない場合はエラー
            if self.config is None:
                logger.error(f"区切りページ生成エラー ({folder_name}): ConfigLoaderが設定されていません")
                return None

            # フォルダ名をセキュアにサニタイズ（PathValidator使用）
            safe_folder_name = PathValidator.sanitize_filename(
                folder_name,
                replacement='_',
                default_name=PDFConversionConstants.DEFAULT_SEPARATOR_NAME
            )

            output_pdf = os.path.join(self.temp_dir, f"separator_{safe_folder_name}.pdf")

            # PDFProcessorで生成（reportlab）
            from pdf_processor import PDFProcessor
            processor = PDFProcessor(self.config)
            return processor.create_separator_pdf(folder_name, output_pdf)

        except (SystemExit, KeyboardInterrupt):
            # システム終了・キーボード割り込みは再スロー
            raise

        except Exception as e:
            logger.exception(f"区切りページ生成エラー ({folder_name}): {e}")
            return None
