"""
PDF変換モジュール（統合版）

各種ファイル（Office、画像、一太郎、PDF）をPDFに変換する機能を提供
このクラスは各変換器のファサードとして機能
"""
import logging
import os
import uuid
from typing import Optional, Any, Callable, Dict, TYPE_CHECKING

from converters.office_converter import OfficeConverter
from converters.image_converter import ImageConverter
from converters.ichitaro_converter import IchitaroConverter
from constants import PDFConversionConstants
from path_validator import PathValidator

if TYPE_CHECKING:
    from config_loader import ConfigLoader
    from pdf_processor import PDFProcessor

logger = logging.getLogger(__name__)


class PDFConverter:
    """各種ファイルをPDFに変換するクラス（ファサード）"""

    # サポートされる拡張子
    OFFICE_EXTENSIONS = OfficeConverter.OFFICE_EXTENSIONS
    IMAGE_EXTENSIONS = ImageConverter.IMAGE_EXTENSIONS
    ICHITARO_EXTENSIONS = IchitaroConverter.ICHITARO_EXTENSIONS

    def __init__(
        self,
        temp_dir: str,
        ichitaro_settings: Optional[Dict[str, Any]] = None,
        cancel_check: Optional[Callable[[], bool]] = None,
        dialog_callback: Optional[Callable[[str, bool], None]] = None,
        config: Optional["ConfigLoader"] = None,
        pdf_processor: Optional["PDFProcessor"] = None
    ) -> None:
        """
        Args:
            temp_dir: 一時ファイルの保存先ディレクトリ
            ichitaro_settings: 一太郎変換のタイミング設定（オプション）
            cancel_check: キャンセル状態をチェックするコールバック関数
            dialog_callback: 一太郎変換ダイアログのコールバック関数(message, show)
            config: ConfigLoaderインスタンス（区切りページ生成に必要）
            pdf_processor: PDFProcessorインスタンス（依存性注入、省略時はconfigから初期化）
        """
        self.temp_dir = temp_dir
        self.config = config

        # 一太郎設定の初期化
        default_ichitaro_settings = {
            'ichitaro_ready_timeout': 30,
            'max_retries': 3,
            'down_arrow_count': 5,
            'save_wait_seconds': 20,
            'dialog_wait_seconds': 3,
            'action_wait_seconds': 2,
        }
        if ichitaro_settings is None:
            self.ichitaro_settings = default_ichitaro_settings.copy()
        else:
            # デフォルト値とマージ
            self.ichitaro_settings = default_ichitaro_settings.copy()
            self.ichitaro_settings.update(ichitaro_settings)

        # PDFProcessorの初期化（依存性注入優先、なければconfigから作成）
        if pdf_processor is not None:
            self._pdf_processor = pdf_processor
        elif config is not None:
            from pdf_processor import PDFProcessor
            self._pdf_processor = PDFProcessor(config)
        else:
            self._pdf_processor = None

        # 各変換器を初期化
        self.office_converter = OfficeConverter(temp_dir)
        self.image_converter = ImageConverter()
        self.ichitaro_converter = IchitaroConverter(
            ichitaro_settings=self.ichitaro_settings,
            cancel_check=cancel_check,
            dialog_callback=dialog_callback
        )

    @staticmethod
    def _is_temporary_file(file_path: str) -> bool:
        """一時ファイルかどうかを判定

        - ~$ を含むファイル名（Office一時ファイル）
        - .$ で始まる拡張子（一太郎一時ファイル: .$td, .$$$ など）
        """
        base_name = os.path.basename(file_path)
        ext = os.path.splitext(file_path)[1].lower()
        # 一時ファイルのパターン: ~$, .$, .$$$など
        return '~$' in base_name or ext.startswith('.$') or base_name.endswith('.$$$')

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
        if self._is_temporary_file(file_path):
            logger.info(f"一時ファイルをスキップ: {os.path.basename(file_path)}")
            return None

        ext = os.path.splitext(file_path)[1].lower()

        # 出力パスの決定（UUID付きで衝突回避）
        if output_path is None:
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            unique_id = uuid.uuid4().hex[:8]
            output_path = os.path.join(self.temp_dir, f"{base_name}_{unique_id}.pdf")

        # 既に変換済みの場合はスキップ
        if os.path.exists(output_path):
            logger.info(f"変換済みファイルが存在: {output_path}")
            return output_path

        # 拡張子に応じて変換
        if ext in self.OFFICE_EXTENSIONS:
            logger.info(f"Officeファイルを変換: {file_path}")
            return self.office_converter.convert(file_path, output_path)
        elif ext == '.pdf':
            logger.debug(f"PDFファイル: {file_path}")
            return file_path if os.path.exists(file_path) else None
        elif ext in self.IMAGE_EXTENSIONS:
            logger.info(f"画像ファイルを変換: {file_path}")
            return self.image_converter.convert(file_path, output_path)
        elif ext in self.ICHITARO_EXTENSIONS:
            logger.info(f"一太郎ファイルを変換: {file_path}")
            return self.ichitaro_converter.convert(file_path, output_path)
        else:
            logger.warning(f"サポートされていないファイル形式: {file_path}")
            return None

    def create_separator_page(self, folder_name: str) -> Optional[str]:
        """
        区切りページを作成（reportlab完全生成版）

        Args:
            folder_name: セクションタイトル

        Returns:
            str: 作成したPDFのパス（失敗時はNone）

        Raises:
            PDFConversionError: ConfigLoaderまたはPDFProcessorが設定されていない場合
        """
        try:
            # ConfigLoaderまたはPDFProcessorが設定されていない場合はエラー
            if self.config is None:
                logger.error(f"区切りページ生成エラー ({folder_name}): ConfigLoaderが設定されていません")
                return None

            if self._pdf_processor is None:
                logger.error(f"区切りページ生成エラー ({folder_name}): PDFProcessorが設定されていません")
                return None

            # フォルダ名をセキュアにサニタイズ（PathValidator使用）
            safe_folder_name = PathValidator.sanitize_filename(
                folder_name,
                replacement='_',
                default_name=PDFConversionConstants.DEFAULT_SEPARATOR_NAME
            )

            output_pdf = os.path.join(self.temp_dir, f"separator_{safe_folder_name}.pdf")

            # PDFProcessorで生成（依存性注入により初期化済み）
            return self._pdf_processor.create_separator_pdf(folder_name, output_pdf)

        except (SystemExit, KeyboardInterrupt):
            raise

        except Exception as e:
            logger.exception(f"区切りページ生成エラー ({folder_name}): {e}")
            return None
