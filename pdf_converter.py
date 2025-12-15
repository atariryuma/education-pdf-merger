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
from typing import Dict, List, Optional, Any, Generator

from win32com import client
import pythoncom
from PIL import Image
from pywinauto import Application
from pywinauto.keyboard import send_keys

# 例外クラスと定数は専用モジュールからインポート
from exceptions import PDFConversionError, CancelledError
from constants import MSOfficeConstants, AppConstants

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
        cancel_check: Optional[callable] = None
    ) -> None:
        """
        Args:
            temp_dir: 一時ファイルの保存先ディレクトリ
            ichitaro_settings: 一太郎変換のタイミング設定（オプション）
            cancel_check: キャンセル状態をチェックするコールバック関数
        """
        self.temp_dir = temp_dir
        # 一太郎設定をマージ（指定されたものでデフォルトを上書き）
        self.ichitaro_settings: Dict[str, Any] = self.DEFAULT_ICHITARO_SETTINGS.copy()
        if ichitaro_settings:
            self.ichitaro_settings.update(ichitaro_settings)
        self._cancel_check = cancel_check or (lambda: False)

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
        interval = 0.5  # 0.5秒間隔でチェック
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
        """Officeファイルを変換"""
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
        """画像ファイルをPDFに変換"""
        image: Optional[Image.Image] = None
        try:
            image = Image.open(file_path)
            if image.mode in ("RGBA", "P"):
                image = image.convert("RGB")
            image.save(output_path, "PDF")
            logger.debug(f"画像変換完了: {file_path} -> {output_path}")
            return output_path
        except IOError as e:
            logger.exception(f"画像変換エラー ({file_path}): {e}")
            raise PDFConversionError(f"画像変換に失敗: {file_path}") from e
        finally:
            if image is not None:
                try:
                    image.close()
                except Exception as e:
                    logger.warning(f"画像ファイルのクローズに失敗: {e}")

    def _convert_ichitaro(self, file_path: str, output_path: str) -> Optional[str]:
        """
        一太郎ファイルをPDFに変換（ベストプラクティス版）

        処理フロー:
        1. 既存の一太郎プロセスをクリーンアップ
        2. 一太郎でファイルを開く
        3. 印刷→保存操作
        4. PDF作成完了を待つ
        5. 一太郎を正常終了
        """
        logger.info("=" * 60)
        logger.info(f"一太郎PDF変換開始: {os.path.basename(file_path)}")
        logger.info("=" * 60)

        # パスを正規化（バックスラッシュに統一）
        file_path = os.path.normpath(file_path)
        output_path = os.path.normpath(output_path)

        logger.info(f"入力ファイル: {file_path}")
        logger.info(f"出力ファイル: {output_path}")

        # 既存の出力ファイルを削除
        if os.path.exists(output_path):
            logger.info(f"既存の出力ファイルを削除: {output_path}")
            os.remove(output_path)

        # 設定値
        max_wait = self.ichitaro_settings.get('ichitaro_ready_timeout', 30)
        down_arrow_count = self.ichitaro_settings.get('down_arrow_count', 5)
        save_wait = self.ichitaro_settings.get('save_wait_seconds', 20)

        logger.info(f"設定値 - 接続タイムアウト: {max_wait}秒, 下矢印回数: {down_arrow_count}, 保存待機: {save_wait}秒")

        app = None
        main_window = None

        try:
            # ステップ1: 事前クリーンアップ
            logger.info("-" * 60)
            logger.info("ステップ1: 事前クリーンアップ")
            logger.info("-" * 60)
            self._cleanup_ichitaro_windows()

            # ステップ2: 一太郎でファイルを開く
            logger.info("-" * 60)
            logger.info("ステップ2: 一太郎でファイルを開く")
            logger.info("-" * 60)
            app, main_window = self._open_ichitaro_file(file_path, max_wait)
            if app is None or main_window is None:
                logger.error("一太郎ファイルを開けませんでした")
                return None

            # ステップ3: 印刷→保存操作
            logger.info("-" * 60)
            logger.info("ステップ3: 印刷→保存操作")
            logger.info("-" * 60)
            self._execute_print_sequence(down_arrow_count, output_path)

            # ステップ4: PDF作成完了を待つ
            logger.info("-" * 60)
            logger.info("ステップ4: PDF作成完了を待つ")
            logger.info("-" * 60)
            result = self._wait_for_output_file(output_path, file_path, save_wait)

            # 最終確認
            if result and os.path.exists(output_path):
                file_size = os.path.getsize(output_path)
                if file_size > 0:
                    logger.info(f"✓ PDFファイル確認OK (サイズ: {file_size:,} bytes)")
                    logger.info("=" * 60)
                    logger.info(f"✓ 一太郎PDF変換成功: {os.path.basename(file_path)}")
                    logger.info("=" * 60)
                else:
                    logger.error(f"✗ PDFファイルが空です (サイズ: 0 bytes)")
                    result = None
            elif result:
                logger.error(f"✗ PDFファイルが削除されています: {output_path}")
                result = None

            if not result:
                logger.error("=" * 60)
                logger.error(f"✗ 一太郎PDF変換失敗: {os.path.basename(file_path)}")
                logger.error("=" * 60)

            return result

        finally:
            # ステップ5: 一太郎を正常終了
            logger.info("-" * 60)
            logger.info("ステップ5: 一太郎を正常終了")
            logger.info("-" * 60)
            self._close_ichitaro(app, main_window)

    def _open_ichitaro_file(self, file_path: str, max_wait: int):
        """
        一太郎でファイルを開く（ベストプラクティス版）

        Args:
            file_path: 一太郎ファイルのパス
            max_wait: 最大待機時間（秒）

        Returns:
            (app, main_window): ApplicationオブジェクトとメインウィンドウのTuple（成功時）
            (None, None): 接続失敗時
        """
        file_name = os.path.basename(file_path)

        if self.is_cancelled():
            raise CancelledError("一太郎変換がキャンセルされました")

        try:
            # ファイルを開く
            logger.info(f"一太郎でファイルを開く: {file_name}")
            os.startfile(file_path)

            # 一太郎起動待機（3秒、キャンセル可能）
            logger.info("一太郎の起動を待機中（3秒）...")
            self._wait_with_cancel_check(3)

            # ファイル名ベースのウィンドウ検索
            import re
            escaped_name = re.escape(file_name)
            title_pattern = f".*{escaped_name}.*"
            logger.info(f"一太郎ウィンドウに接続中: '{title_pattern}'")

            # pywinautoで接続（タイムアウト付き）
            app = Application(backend="uia").connect(
                title_re=title_pattern, timeout=max_wait
            )
            main_window = app.top_window()
            logger.info(f"✓ 一太郎への接続成功: {main_window.window_text()}")
            main_window.set_focus()
            return app, main_window

        except CancelledError:
            self._cleanup_ichitaro_windows()
            raise
        except Exception as e:
            logger.error(f"✗ 一太郎への接続失敗: {e}")
            self._cleanup_ichitaro_windows()
            return None, None

    def _close_ichitaro(self, app, main_window):
        """
        一太郎を正常終了（ベストプラクティス版）

        pywinautoの推奨方法で一太郎を閉じる：
        1. window.close() で正常終了を試みる
        2. 失敗した場合のみ app.kill() で強制終了
        """
        if app is None or main_window is None:
            logger.info("一太郎が開いていないため、クローズ不要")
            return

        try:
            # 印刷処理完了を待つ
            logger.info("印刷処理の完全終了を待機中（2秒）...")
            self._wait_with_cancel_check(2)

            # pywinauto推奨: window.close()で正常終了
            logger.info("一太郎ウィンドウを閉じています（window.close()）...")
            window_title = main_window.window_text()
            logger.info(f"対象ウィンドウ: {window_title}")

            main_window.close()
            logger.info("✓ 一太郎ウィンドウを閉じました")

            # ウィンドウが実際に閉じたことを確認
            self._wait_with_cancel_check(1)

        except Exception as e:
            logger.warning(f"window.close()が失敗しました: {e}")
            logger.info("強制終了を試みます（app.kill()）...")
            try:
                app.kill()
                logger.info("✓ 一太郎を強制終了しました")
            except Exception as e2:
                logger.error(f"強制終了も失敗: {e2}")

    def _execute_print_sequence(self, down_arrow_count: int, output_path: str) -> bool:
        """
        印刷シーケンスを実行（改善版）

        手順：
        Ctrl+P → 下矢印でプリンタ選択 → Enter → ファイル名入力 → Enter

        Returns:
            bool: 処理成功時True、失敗時False
        """
        # Ctrl+P で印刷ダイアログを開く
        logger.info("印刷ダイアログを開く (Ctrl+P)")
        send_keys("^p")
        logger.info("Ctrl+P送信完了、2秒待機...")
        self._wait_with_cancel_check(2)
        logger.info("✓ 2秒待機完了、印刷ダイアログが開いているはず")

        # 下矢印でMicrosoft Print to PDFを選択
        logger.info(f"プリンタを選択 (↓×{down_arrow_count})")
        for i in range(down_arrow_count):
            if self.is_cancelled():
                raise CancelledError("一太郎変換がキャンセルされました")
            send_keys("{DOWN}")
            logger.debug(f"  下矢印 {i+1}/{down_arrow_count}")
            self._wait_with_cancel_check(0.1)
        logger.info("プリンタ選択完了、0.5秒待機...")
        self._wait_with_cancel_check(0.5)

        if self.is_cancelled():
            raise CancelledError("一太郎変換がキャンセルされました")

        # Enter で印刷実行（2回押す）
        logger.info("=" * 60)
        logger.info("印刷ダイアログでEnterキーを2回押します")
        logger.info("=" * 60)
        logger.info("1回目のEnter...")
        send_keys("{ENTER}")
        self._wait_with_cancel_check(0.3)
        logger.info("2回目のEnter...")
        send_keys("{ENTER}")
        logger.info("✓ Enterキー2回送信完了")
        logger.info("保存ダイアログの表示を待機中（2秒）...")
        self._wait_with_cancel_check(2)  # 保存ダイアログの表示を待つ
        logger.info("✓ 2秒待機完了")

        # ダイアログ検出をスキップして、直接キーボード操作でファイル名入力
        logger.info("=" * 60)
        logger.info("保存ダイアログへキーボード操作で直接入力")
        logger.info("=" * 60)

        logger.info("Ctrl+Aで全選択...")
        send_keys("^a")
        self._wait_with_cancel_check(0.3)

        logger.info(f"ファイルパスを入力: {output_path}")
        escaped_path = self._escape_for_send_keys(output_path)
        logger.info(f"エスケープ済みパス: {escaped_path}")
        send_keys(escaped_path, pause=0.02, with_spaces=True)
        self._wait_with_cancel_check(0.5)

        logger.info("Enterキーで保存実行...")
        send_keys("{ENTER}")
        logger.info("✓ 保存処理完了")
        return True

    def _wait_for_output_file(self, output_path: str, file_path: str, save_wait: int = 20) -> Optional[str]:
        """出力ファイルの作成を待機（動的間隔）"""
        logger.info(f"出力ファイルの作成を待機中（最大{save_wait}秒、動的間隔でチェック）...")
        logger.info(f"待機対象ファイル: {output_path}")

        # 動的待機間隔: 最初は0.1秒、その後0.5秒、最後は1秒
        wait_intervals = [0.1] * 10 + [0.5] * 20 + [1.0] * 20  # 最大51秒

        elapsed_time = 0.0
        last_size = 0
        stable_count = 0

        for i, interval in enumerate(wait_intervals):
            if elapsed_time > save_wait:
                break

            if os.path.exists(output_path):
                current_size = os.path.getsize(output_path)

                # ファイルサイズが0より大きく、安定している（3回連続で同じサイズ）
                if current_size > 0:
                    if current_size == last_size:
                        stable_count += 1
                        if stable_count >= 3:  # 1.5秒間サイズ不変なら完成
                            logger.info(f"✓ 出力ファイル検出成功！ (サイズ: {current_size:,} bytes)")
                            logger.info(f"待機時間: {elapsed_time:.1f}秒")
                            return output_path
                    else:
                        stable_count = 0
                    last_size = current_size

            # 5秒ごとに経過時間をログ出力
            if elapsed_time > 0 and int(elapsed_time) % 5 == 0 and i > 0:
                logger.info(f"待機中... 経過時間: {elapsed_time:.1f}秒 / {save_wait}秒")

            self._wait_with_cancel_check(interval)
            elapsed_time += interval

        logger.error(f"✗ タイムアウト: {save_wait}秒経過しても出力ファイルが見つかりません")
        logger.error(f"ファイルパス: {output_path}")
        return None

    def _cleanup_ichitaro_windows(self) -> None:
        """残っている一太郎ウィンドウをクリーンアップ（最適化版）"""
        try:
            logger.info("一太郎ウィンドウのクリーンアップを開始...")
            app = Application(backend="uia").connect(title_re=".*一太郎.*", timeout=1)

            # pywinauto推奨: app.kill()で全プロセス終了
            logger.info("一太郎プロセスを強制終了しています...")
            app.kill()
            self._wait_with_cancel_check(0.5)
            logger.info("✓ 一太郎プロセスのクリーンアップ完了")

        except Exception:
            # 一太郎が起動していない場合は正常
            logger.info("一太郎プロセスなし（クリーンアップ不要）")

    def create_separator_page(self, folder_name: str, template_path: str) -> Optional[str]:
        """
        区切りページ（中扉）を作成

        Args:
            folder_name: フォルダ名（ページに挿入するテキスト）
            template_path: Wordテンプレートのパス

        Returns:
            str: 作成したPDFのパス（失敗時はNone）

        Raises:
            PDFConversionError: 区切りページ作成中にエラーが発生した場合
        """
        # フォルダ名をサニタイズ
        safe_folder_name = re.sub(r'[<>:"/\\|?*]', '_', folder_name)
        output_pdf = os.path.join(self.temp_dir, f"separator_{safe_folder_name}.pdf")

        if os.path.exists(output_pdf):
            try:
                os.remove(output_pdf)
            except OSError as e:
                logger.warning(f"既存の区切りページの削除に失敗: {e}")

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
                doc = word.Documents.Open(template_path, ReadOnly=True)
                doc.Content.InsertBefore(folder_name + "\n")
                doc.SaveAs2(output_pdf, FileFormat=self.WORD_PDF_FORMAT)

                logger.info(f"区切りページを作成: {folder_name}")
                return output_pdf
            except Exception as e:
                logger.exception(f"区切りページ作成エラー ({folder_name}): {e}")
                raise PDFConversionError(f"区切りページ作成に失敗: {folder_name}") from e
            finally:
                self._cleanup_office_app(doc, word, "WINWORD.EXE", "Word")
