"""
一太郎変換モジュール

一太郎ファイル(.jtd)をPDFに変換する機能を提供
"""
import logging
import os
import re
import time
from typing import Optional, Callable, Dict, Any, Tuple

from pywinauto import Application
from pywinauto.keyboard import send_keys

from exceptions import CancelledError
from constants import AppConstants, PDFConversionConstants, IchitaroWaitTimes

logger = logging.getLogger(__name__)


class IchitaroConverter:
    """一太郎ファイルをPDFに変換するクラス"""

    ICHITARO_EXTENSIONS = ('.jtd',)

    def __init__(
        self,
        ichitaro_settings: Optional[Dict[str, Any]] = None,
        cancel_check: Optional[Callable[[], bool]] = None,
        dialog_callback: Optional[Callable[[str, bool], None]] = None
    ) -> None:
        """
        Args:
            ichitaro_settings: 一太郎変換のタイミング設定（オプション）
            cancel_check: キャンセル状態をチェックするコールバック関数
            dialog_callback: 一太郎変換ダイアログのコールバック関数(message, show)
        """
        self.ichitaro_settings = ichitaro_settings or {
            **AppConstants.ICHITARO_DEFAULTS,
            'printer_name': 'Microsoft Print to PDF'
        }
        self._cancel_check = cancel_check or (lambda: False)
        self._dialog_callback = dialog_callback

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
    def _escape_for_send_keys(text: str) -> str:
        """
        send_keys用に特殊文字をエスケープ

        Args:
            text: エスケープする文字列

        Returns:
            str: エスケープされた文字列

        Note:
            pywinauto公式仕様のエスケープルール:
            - { → {{
            - } → }}
            - 他の特殊文字（+, ^, %, ~, (, )）→ {文字}
        """
        # { と } は二重にしてエスケープ（最初に処理）
        text = text.replace('{', '{{').replace('}', '}}')

        # 他の特殊文字は {文字} でエスケープ
        special_chars = {'+': '{+}', '^': '{^}', '%': '{%}', '~': '{~}', '(': '{(}', ')': '{)}'}
        for char, escaped in special_chars.items():
            text = text.replace(char, escaped)

        return text

    def convert(self, file_path: str, output_path: str) -> Optional[str]:
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
        max_attempts = IchitaroWaitTimes.MAX_ATTEMPTS
        file_name = os.path.basename(file_path)

        # ダイアログ表示（ファイル変換の最初に1回だけ）
        if self._dialog_callback:
            self._dialog_callback(f"変換中: {file_name}", True)  # True = show

        try:
            for attempt in range(1, max_attempts + 1):
                if attempt > 1:
                    logger.info(f"一太郎PDF変換を再試行します（試行 {attempt}/{max_attempts}）")
                logger.info(f"一太郎PDF変換開始: {file_name}")

                # パスを正規化（バックスラッシュに統一）
                file_path_norm = os.path.normpath(file_path)
                output_path_norm = os.path.normpath(output_path)

                logger.info(f"入力ファイル: {file_path_norm}")
                logger.info(f"出力ファイル: {output_path_norm}")

                # 既存の出力ファイルを削除（TOCTOU脆弱性回避）
                try:
                    os.remove(output_path_norm)
                    logger.info(f"既存の出力ファイルを削除: {output_path_norm}")
                except FileNotFoundError:
                    pass
                except Exception as e:
                    logger.warning(f"出力ファイル削除エラー（続行します）: {e}")

                # 設定値
                max_wait = self.ichitaro_settings.get('ichitaro_ready_timeout', 30)
                save_wait = self.ichitaro_settings.get('save_wait_seconds', 20)

                logger.info(f"設定値 - 接続タイムアウト: {max_wait}秒, 保存待機: {save_wait}秒")

                app = None
                main_window = None

                try:
                    # ステップ1: 事前クリーンアップ
                    logger.debug("ステップ1: 事前クリーンアップ")
                    self._cleanup_ichitaro_windows()

                    # ステップ2: 一太郎でファイルを開く
                    logger.debug("ステップ2: 一太郎でファイルを開く")
                    app, main_window = self._open_ichitaro_file(file_path_norm, max_wait)
                    if app is None or main_window is None:
                        logger.error("一太郎ファイルを開けませんでした")
                        result = None
                    else:
                        # ステップ3: 印刷→保存操作
                        logger.debug("ステップ3: 印刷→保存操作")
                        self._execute_print_sequence(app, output_path_norm)

                        # ステップ4: PDF作成完了を待つ
                        logger.debug("ステップ4: PDF作成完了を待つ")
                        result = self._wait_for_output_file(output_path_norm, file_path_norm, save_wait)

                finally:
                    # ステップ5: 一太郎を正常終了
                    logger.debug("ステップ5: 一太郎を正常終了")
                    self._close_ichitaro(app, main_window)

                # 最終確認
                if result and os.path.exists(output_path_norm):
                    file_size = os.path.getsize(output_path_norm)
                    if file_size > 0:
                        logger.debug(f"PDFファイル確認OK (サイズ: {file_size:,} bytes)")
                        if attempt > 1:
                            logger.info(f"一太郎PDF変換成功（試行 {attempt}/{max_attempts}）: {file_name}")
                        else:
                            logger.info(f"一太郎PDF変換成功: {file_name}")
                        return result
                    else:
                        logger.error(f"{PDFConversionConstants.LOG_MARK_FAILURE} PDFファイルが空です (サイズ: 0 bytes)")
                        result = None
                elif result:
                    logger.error(f"{PDFConversionConstants.LOG_MARK_FAILURE} PDFファイルが削除されています: {output_path_norm}")
                    result = None

                # 変換失敗時の処理
                logger.error(f"一太郎PDF変換失敗（試行 {attempt}/{max_attempts}）: {file_name}")

                if attempt < max_attempts:
                    logger.warning("一太郎プロセスをクリーンアップして再試行します...")
                    self._cleanup_ichitaro_windows()
                    self._wait_with_cancel_check(IchitaroWaitTimes.RETRY_DELAY)
                    # ダイアログメッセージ更新
                    if self._dialog_callback:
                        self._dialog_callback(f"再試行中: {file_name} ({attempt+1}/{max_attempts})", True)
                else:
                    logger.error(f"{PDFConversionConstants.LOG_MARK_FAILURE} 一太郎変換が{max_attempts}回失敗しました。スキップします: {file_name}")

            # ループを抜けた = 全試行失敗
            logger.warning(f"一太郎変換が全{max_attempts}回の試行で失敗しました: {file_name}")
            return None

        except CancelledError:
            logger.info("一太郎変換がキャンセルされました")
            raise

        except (SystemExit, KeyboardInterrupt):
            raise

        except Exception as e:
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
            wait_time = IchitaroWaitTimes.STARTUP_WAIT
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
            wait_time = IchitaroWaitTimes.PRINT_COMPLETE_WAIT
            logger.info(f"印刷処理の完全終了を待機中（{wait_time}秒）...")
            self._wait_with_cancel_check(wait_time)

            # 一太郎を強制終了（保存確認ダイアログを回避）
            logger.info("一太郎プロセスを終了中（app.kill()）...")
            window_title = main_window.window_text()
            logger.info(f"対象ウィンドウ: {window_title}")

            app.kill()
            logger.info(f"{PDFConversionConstants.LOG_MARK_SUCCESS} 一太郎プロセスを終了しました")

            # プロセス終了の確認待機
            wait_time = IchitaroWaitTimes.WINDOW_CLOSE_WAIT
            self._wait_with_cancel_check(wait_time)

        except Exception as e:
            logger.warning(f"一太郎の終了に失敗しました: {e}")

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
        # Ctrl+P で印刷ダイアログを開く
        logger.info("印刷ダイアログを開く (Ctrl+P)")
        send_keys("^p")
        logger.info("Ctrl+P送信完了、印刷ダイアログの表示を待機中...")
        wait_time = IchitaroWaitTimes.CTRL_P_WAIT
        self._wait_with_cancel_check(wait_time)
        logger.info(f"{PDFConversionConstants.LOG_MARK_SUCCESS} 待機完了、印刷ダイアログが開いているはず")

        # Microsoft Print to PDFをプリンター名で直接選択
        logger.info("Microsoft Print to PDFを選択中...")

        # リトライ機構：低スペックPCでダイアログの準備が遅い場合に対応
        max_retries = PDFConversionConstants.PRINTER_SELECT_MAX_RETRIES
        retry_delay = PDFConversionConstants.PRINTER_SELECT_RETRY_DELAY

        for attempt in range(max_retries):
            try:
                main_window = app.top_window()
                print_dialog = main_window.child_window(title="印刷", control_type="Window")
                printer_combo = print_dialog.child_window(auto_id="1297", control_type="ComboBox")

                # pywinautoの高レベルAPIでプリンターを選択
                printer_combo.select("Microsoft Print to PDF")
                logger.info(f"{PDFConversionConstants.LOG_MARK_SUCCESS} select()メソッドで'Microsoft Print to PDF'を選択")
                self._wait_with_cancel_check(IchitaroWaitTimes.PRINTER_SELECT_WAIT)

                # 印刷ボタンにフォーカスを設定
                try:
                    ok_button = print_dialog.child_window(title="OK", control_type="Button")
                    ok_button.set_focus()
                    logger.info("印刷ボタン（OK）にフォーカスを設定")
                    self._wait_with_cancel_check(IchitaroWaitTimes.CTRL_A_WAIT)
                except Exception as focus_error:
                    logger.debug(f"印刷ボタンへのフォーカス設定をスキップ: {focus_error}")

                break

            except Exception as select_error:
                if attempt < max_retries - 1:
                    logger.warning(f"プリンター選択失敗（試行 {attempt + 1}/{max_retries}）: {select_error}")
                    logger.info(f"{retry_delay}秒待機後、再試行します...")
                    self._wait_with_cancel_check(retry_delay)
                else:
                    logger.error(f"プリンター選択が{max_retries}回失敗しました: {select_error}")
                    raise Exception(f"Microsoft Print to PDFの選択に失敗しました: {select_error}")

        if self.is_cancelled():
            raise CancelledError("一太郎変換がキャンセルされました")

        # Enter で印刷実行（2回押す）
        logger.debug("印刷ダイアログでEnterキーを2回押します")
        send_keys("{ENTER}")
        self._wait_with_cancel_check(IchitaroWaitTimes.ENTER_INTERVAL)
        send_keys("{ENTER}")
        logger.debug("Enterキー2回送信完了")

        # 保存ダイアログ検出と入力
        self._handle_save_dialog(app, output_path)
        return True

    def _try_detect_save_dialog(self, app: Any, dialog_elapsed: float) -> bool:
        """
        保存ダイアログの検出を試行（3つの方法で試行）

        Args:
            app: pywinauto Application
            dialog_elapsed: 経過時間

        Returns:
            bool: ダイアログが検出できたか
        """
        try:
            # 方法1: top_window()で最前面ウィンドウを取得
            try:
                top = app.top_window()
                if top and top.exists(timeout=0):
                    class_name = top.class_name()
                    if class_name == "#32770" or "JSTARO" not in class_name.upper():
                        logger.info(f"{PDFConversionConstants.LOG_MARK_SUCCESS} 保存ダイアログ検出（top_window, {dialog_elapsed:.1f}秒経過）")
                        logger.info(f"  クラス名: {class_name}")
                        return True
            except Exception as e:
                logger.debug(f"top_window検出エラー: {e}")

            # 方法2: タイトル正規表現での検出
            save_dialogs = app.windows(title_re='.*名前を付けて保存.*|.*Save.*|.*保存.*')
            if save_dialogs:
                for dlg in save_dialogs:
                    if dlg.exists(timeout=0):
                        logger.info(f"{PDFConversionConstants.LOG_MARK_SUCCESS} 保存ダイアログ検出（title_re, {dialog_elapsed:.1f}秒経過）")
                        return True

            # 方法3: クラス名#32770での検出
            try:
                dialog_windows = app.windows(class_name="#32770")
                if dialog_windows:
                    for dlg in dialog_windows:
                        if dlg.exists(timeout=0):
                            logger.info(f"{PDFConversionConstants.LOG_MARK_SUCCESS} 保存ダイアログ検出（#32770, {dialog_elapsed:.1f}秒経過）")
                            logger.info(f"  タイトル: {dlg.window_text()}")
                            return True
            except Exception as e:
                logger.debug(f"#32770検出試行エラー: {e}")

        except Exception as e:
            logger.debug(f"ダイアログ検出試行エラー: {e}")

        return False

    def _handle_save_dialog(self, app: Any, output_path: str) -> None:
        """
        保存ダイアログの表示を待機し、キーボード操作で入力

        Args:
            app: pywinauto Application
            output_path: 出力ファイルパス
        """
        # 保存ダイアログの表示を動的に待機
        logger.debug(f"保存ダイアログの表示を待機中（最大{IchitaroWaitTimes.DIALOG_TIMEOUT}秒）")

        dialog_timeout = IchitaroWaitTimes.DIALOG_TIMEOUT
        dialog_wait_interval = IchitaroWaitTimes.DIALOG_POLL_INTERVAL
        dialog_elapsed = 0.0
        dialog_found = False
        min_wait = IchitaroWaitTimes.DIALOG_MIN_WAIT

        while dialog_elapsed < dialog_timeout:
            if self.is_cancelled():
                logger.info("一太郎変換がキャンセルされました（ダイアログ待機中）")
                self._cleanup_ichitaro_windows()
                raise CancelledError("一太郎変換がキャンセルされました")

            self._wait_with_cancel_check(dialog_wait_interval)
            dialog_elapsed += dialog_wait_interval

            # 最低待機時間経過後、ダイアログ検出を試行
            if dialog_elapsed >= min_wait:
                dialog_found = self._try_detect_save_dialog(app, dialog_elapsed)
                if dialog_found:
                    break

            # 進行状況を表示（2秒ごと）
            if dialog_elapsed >= 4.0 and dialog_elapsed % 2.0 < dialog_wait_interval:
                logger.info(f"待機中... {dialog_elapsed:.1f}秒 / {dialog_timeout}秒")

        if dialog_found:
            logger.info(f"{PDFConversionConstants.LOG_MARK_SUCCESS} 保存ダイアログ確認済み")
        else:
            logger.info(f"{PDFConversionConstants.LOG_MARK_SUCCESS} {dialog_elapsed:.1f}秒待機完了（ダイアログ検出なし、キーボード操作で続行）")

        # キーボード入力の準備のため追加待機
        self._wait_with_cancel_check(IchitaroWaitTimes.KEYBOARD_PREP_WAIT)

        # キーボード操作でファイル名入力
        logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)
        logger.info("保存ダイアログへキーボード操作で直接入力")
        logger.info(PDFConversionConstants.LOG_SEPARATOR_MAJOR)

        logger.info("Ctrl+Aで全選択...")
        send_keys("^a")
        self._wait_with_cancel_check(IchitaroWaitTimes.CTRL_A_WAIT)

        logger.info(f"ファイルパスを入力: {output_path}")
        escaped_path = self._escape_for_send_keys(output_path)
        logger.info(f"エスケープ済みパス: {escaped_path}")
        send_keys(escaped_path, pause=0.02, with_spaces=True)
        self._wait_with_cancel_check(IchitaroWaitTimes.FILE_INPUT_WAIT)

        logger.info("Enterキーで保存実行...")
        send_keys("{ENTER}")
        logger.info(f"{PDFConversionConstants.LOG_MARK_SUCCESS} 保存処理完了")

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

        # 動的待機間隔を生成
        def generate_intervals():
            """待機間隔を生成するジェネレータ"""
            for _ in range(PDFConversionConstants.FILE_WAIT_FAST_COUNT):
                yield PDFConversionConstants.FILE_WAIT_INTERVAL_FAST
            for _ in range(PDFConversionConstants.FILE_WAIT_MEDIUM_COUNT):
                yield PDFConversionConstants.FILE_WAIT_INTERVAL_MEDIUM
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

                # ファイルサイズが0より大きく、安定している
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
            timeout = IchitaroWaitTimes.CLEANUP_TIMEOUT
            app = Application(backend="uia").connect(title_re=".*一太郎.*", timeout=timeout)

            logger.info("一太郎プロセスを強制終了しています...")
            app.kill()
            self._wait_with_cancel_check(IchitaroWaitTimes.CLEANUP_WAIT)
            logger.info(f"{PDFConversionConstants.LOG_MARK_SUCCESS} 一太郎プロセスのクリーンアップ完了")

        except Exception:
            logger.info("一太郎プロセスなし（クリーンアップ不要）")
