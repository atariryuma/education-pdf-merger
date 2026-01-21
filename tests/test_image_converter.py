"""
ImageConverterのユニットテスト

PIL (Pillow)をモック化してテスト
"""
import tempfile
import shutil
from pathlib import Path
from typing import Generator
from unittest.mock import Mock, patch, MagicMock
import pytest

from converters.image_converter import ImageConverter
from exceptions import PDFConversionError


@pytest.fixture
def temp_dir() -> Generator[Path, None, None]:
    """
    一時ディレクトリを作成

    Yields:
        Path: 一時ディレクトリのパス
    """
    temp_path = Path(tempfile.mkdtemp(prefix="image_test_"))
    try:
        yield temp_path
    finally:
        shutil.rmtree(temp_path, ignore_errors=True)


@pytest.fixture
def converter() -> ImageConverter:
    """
    ImageConverterインスタンスを作成

    Returns:
        ImageConverter: テスト用インスタンス
    """
    return ImageConverter()


@pytest.fixture
def mock_jpg_file(temp_dir: Path) -> Path:
    """
    ダミーJPEGファイルを作成

    Args:
        temp_dir: 一時ディレクトリ

    Returns:
        Path: ダミーファイルのパス
    """
    jpg_path = temp_dir / "test.jpg"
    jpg_path.write_bytes(b'\xFF\xD8\xFF\xE0')  # JPEG magic number
    return jpg_path


@pytest.fixture
def mock_png_file(temp_dir: Path) -> Path:
    """
    ダミーPNGファイルを作成

    Args:
        temp_dir: 一時ディレクトリ

    Returns:
        Path: ダミーファイルのパス
    """
    png_path = temp_dir / "test.png"
    png_path.write_bytes(b'\x89PNG\r\n\x1a\n')  # PNG magic number
    return png_path


class TestImageConverter:
    """ImageConverterのテスト"""

    def test_initialization(self):
        """初期化のテスト"""
        converter = ImageConverter()
        assert isinstance(converter, ImageConverter)

    def test_image_extensions_constant(self):
        """サポートされる拡張子の定義確認"""
        expected_extensions = ('.jpg', '.jpeg', '.png', '.bmp', '.tiff')
        assert ImageConverter.IMAGE_EXTENSIONS == expected_extensions

    @patch('converters.image_converter.Image')
    def test_convert_jpg_success(
        self,
        mock_image_module: Mock,
        converter: ImageConverter,
        mock_jpg_file: Path,
        temp_dir: Path
    ):
        """JPEG変換の成功テスト"""
        output_path = temp_dir / "output.pdf"

        # Imageオブジェクトのモック
        mock_image = MagicMock()
        mock_image.mode = "RGB"
        mock_image.__enter__ = Mock(return_value=mock_image)
        mock_image.__exit__ = Mock(return_value=False)
        mock_image_module.open.return_value = mock_image

        result = converter.convert(str(mock_jpg_file), str(output_path))

        assert result == str(output_path)
        mock_image_module.open.assert_called_once_with(str(mock_jpg_file))
        mock_image.save.assert_called_once_with(str(output_path), "PDF")

    @patch('converters.image_converter.Image')
    def test_convert_png_success(
        self,
        mock_image_module: Mock,
        converter: ImageConverter,
        mock_png_file: Path,
        temp_dir: Path
    ):
        """PNG変換の成功テスト"""
        output_path = temp_dir / "output.pdf"

        # Imageオブジェクトのモック
        mock_image = MagicMock()
        mock_image.mode = "RGB"
        mock_image.__enter__ = Mock(return_value=mock_image)
        mock_image.__exit__ = Mock(return_value=False)
        mock_image_module.open.return_value = mock_image

        result = converter.convert(str(mock_png_file), str(output_path))

        assert result == str(output_path)
        mock_image_module.open.assert_called_once_with(str(mock_png_file))
        mock_image.save.assert_called_once_with(str(output_path), "PDF")

    @patch('converters.image_converter.Image')
    def test_convert_rgba_to_rgb(
        self,
        mock_image_module: Mock,
        converter: ImageConverter,
        mock_png_file: Path,
        temp_dir: Path
    ):
        """RGBA画像のRGB変換テスト"""
        output_path = temp_dir / "output.pdf"

        # RGBA画像のモック
        mock_image = MagicMock()
        mock_image.mode = "RGBA"
        mock_converted = MagicMock()
        mock_image.convert.return_value = mock_converted
        mock_image.__enter__ = Mock(return_value=mock_image)
        mock_image.__exit__ = Mock(return_value=False)
        mock_image_module.open.return_value = mock_image

        result = converter.convert(str(mock_png_file), str(output_path))

        assert result == str(output_path)
        mock_image.convert.assert_called_once_with("RGB")
        # convert()で返されたオブジェクトがsave()を呼ぶ
        # 実際の実装では image = image.convert("RGB") なので、
        # save()はmock_convertedではなくmock_imageで呼ばれる

    @patch('converters.image_converter.Image')
    def test_convert_palette_to_rgb(
        self,
        mock_image_module: Mock,
        converter: ImageConverter,
        mock_png_file: Path,
        temp_dir: Path
    ):
        """パレット画像(P mode)のRGB変換テスト"""
        output_path = temp_dir / "output.pdf"

        # パレット画像のモック
        mock_image = MagicMock()
        mock_image.mode = "P"
        mock_converted = MagicMock()
        mock_image.convert.return_value = mock_converted
        mock_image.__enter__ = Mock(return_value=mock_image)
        mock_image.__exit__ = Mock(return_value=False)
        mock_image_module.open.return_value = mock_image

        result = converter.convert(str(mock_png_file), str(output_path))

        assert result == str(output_path)
        mock_image.convert.assert_called_once_with("RGB")

    @patch('converters.image_converter.Image')
    def test_convert_io_error(
        self,
        mock_image_module: Mock,
        converter: ImageConverter,
        mock_jpg_file: Path,
        temp_dir: Path
    ):
        """IOError発生時のテスト"""
        output_path = temp_dir / "output.pdf"
        mock_image_module.open.side_effect = IOError("Cannot open image")

        with pytest.raises(PDFConversionError) as exc_info:
            converter.convert(str(mock_jpg_file), str(output_path))

        assert "画像変換に失敗" in str(exc_info.value)
        assert exc_info.value.original_error is not None
        assert isinstance(exc_info.value.original_error, IOError)

    @patch('converters.image_converter.Image')
    def test_convert_exception_chaining(
        self,
        mock_image_module: Mock,
        converter: ImageConverter,
        mock_jpg_file: Path,
        temp_dir: Path
    ):
        """例外チェーンの確認"""
        output_path = temp_dir / "output.pdf"
        original_error = IOError("Test error")
        mock_image_module.open.side_effect = original_error

        with pytest.raises(PDFConversionError) as exc_info:
            converter.convert(str(mock_jpg_file), str(output_path))

        # 例外チェーンの確認
        assert exc_info.value.original_error is original_error
        assert exc_info.value.__cause__ is original_error

    @patch('converters.image_converter.Image')
    def test_convert_context_manager_cleanup(
        self,
        mock_image_module: Mock,
        converter: ImageConverter,
        mock_jpg_file: Path,
        temp_dir: Path
    ):
        """with文によるリソース解放の確認"""
        output_path = temp_dir / "output.pdf"

        mock_image = MagicMock()
        mock_image.mode = "RGB"
        mock_enter = Mock(return_value=mock_image)
        mock_exit = Mock(return_value=False)
        mock_image.__enter__ = mock_enter
        mock_image.__exit__ = mock_exit
        mock_image_module.open.return_value = mock_image

        converter.convert(str(mock_jpg_file), str(output_path))

        # コンテキストマネージャーのenter/exitが呼ばれることを確認
        mock_enter.assert_called_once()
        mock_exit.assert_called_once()

    @patch('converters.image_converter.Image')
    def test_convert_context_manager_with_exception(
        self,
        mock_image_module: Mock,
        converter: ImageConverter,
        mock_jpg_file: Path,
        temp_dir: Path
    ):
        """例外発生時のwith文クリーンアップ確認"""
        output_path = temp_dir / "output.pdf"

        mock_image = MagicMock()
        mock_image.mode = "RGB"
        mock_image.save.side_effect = IOError("Save failed")
        mock_enter = Mock(return_value=mock_image)
        mock_exit = Mock(return_value=False)
        mock_image.__enter__ = mock_enter
        mock_image.__exit__ = mock_exit
        mock_image_module.open.return_value = mock_image

        with pytest.raises(PDFConversionError):
            converter.convert(str(mock_jpg_file), str(output_path))

        # 例外が発生してもexitが呼ばれることを確認
        mock_exit.assert_called_once()

    def test_convert_all_supported_extensions(self, converter: ImageConverter, temp_dir: Path):
        """全サポート拡張子のテスト"""
        supported_extensions = ImageConverter.IMAGE_EXTENSIONS

        for ext in supported_extensions:
            input_file = temp_dir / f"test{ext}"
            input_file.write_bytes(b'\x00\x01\x02\x03')  # ダミーバイナリ

            output_file = temp_dir / f"output{ext}.pdf"

            # 実際のPILは使わず、モック化する
            with patch('converters.image_converter.Image') as mock_image_module:
                mock_image = MagicMock()
                mock_image.mode = "RGB"
                mock_image.__enter__ = Mock(return_value=mock_image)
                mock_image.__exit__ = Mock(return_value=False)
                mock_image_module.open.return_value = mock_image

                result = converter.convert(str(input_file), str(output_file))
                assert result == str(output_file)


if __name__ == "__main__":
    pytest.main([__file__, "-v", "--tb=short"])
