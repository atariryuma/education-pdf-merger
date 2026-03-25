"""
画像変換モジュール

画像ファイル（JPEG、PNG、BMP、TIFF）をPDFに変換する機能を提供
"""
import logging
from typing import Optional

from PIL import Image

from shared.exceptions import PDFConversionError

logger = logging.getLogger(__name__)


class ImageConverter:
    """画像ファイルをPDFに変換するクラス"""

    IMAGE_EXTENSIONS = ('.jpg', '.jpeg', '.png', '.bmp', '.tiff')

    def convert(self, file_path: str, output_path: str) -> Optional[str]:
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
                # RGB以外のモード（RGBA, P, CMYK, LAなど）はRGBに変換
                if image.mode != "RGB":
                    rgb_image = image.convert("RGB")
                    try:
                        rgb_image.save(output_path, "PDF")
                    finally:
                        rgb_image.close()
                else:
                    image.save(output_path, "PDF")

            logger.debug(f"画像変換完了: {file_path} -> {output_path}")
            return output_path
        except (IOError, ValueError, SyntaxError) as e:
            logger.error(f"画像変換エラー ({file_path}): {e}")
            raise PDFConversionError(f"画像変換に失敗: {file_path}", original_error=e) from e
