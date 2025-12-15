"""
PDFConverterのテスト
"""
import os
import pytest
import sys

# プロジェクトルートをパスに追加
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pdf_converter import PDFConverter


class TestPDFConverter:
    """PDFConverterクラスのテスト"""

    def test_init_with_default_settings(self, temp_dir):
        """デフォルト設定での初期化"""
        converter = PDFConverter(temp_dir)
        assert converter.temp_dir == temp_dir
        assert converter.ichitaro_settings['ichitaro_ready_timeout'] == 30
        assert converter.ichitaro_settings['max_retries'] == 3
        assert converter.ichitaro_settings['down_arrow_count'] == 5
        assert converter.ichitaro_settings['save_wait_seconds'] == 20

    def test_init_with_custom_settings(self, temp_dir):
        """カスタム設定での初期化"""
        custom_settings = {
            'ichitaro_ready_timeout': 60,
            'max_retries': 5,
            'down_arrow_count': 7,
            'save_wait_seconds': 30
        }
        converter = PDFConverter(temp_dir, custom_settings)
        assert converter.ichitaro_settings['ichitaro_ready_timeout'] == 60
        assert converter.ichitaro_settings['max_retries'] == 5
        assert converter.ichitaro_settings['down_arrow_count'] == 7
        assert converter.ichitaro_settings['save_wait_seconds'] == 30

    def test_skip_temporary_file(self, temp_dir):
        """一時ファイルのスキップ"""
        converter = PDFConverter(temp_dir)
        # ~$で始まるファイルはスキップ
        result = converter.convert(os.path.join(temp_dir, "~$temp.docx"))
        assert result is None

    def test_convert_unsupported_file(self, temp_dir):
        """サポートされていないファイル形式"""
        converter = PDFConverter(temp_dir)
        # 存在するテキストファイルを作成
        txt_path = os.path.join(temp_dir, "test.txt")
        with open(txt_path, 'w') as f:
            f.write("test content")

        result = converter.convert(txt_path)
        assert result is None

    def test_convert_already_pdf(self, temp_dir, sample_pdf_file):
        """既存PDFファイルの素通し"""
        converter = PDFConverter(temp_dir)
        result = converter.convert(sample_pdf_file)
        assert result == sample_pdf_file

    def test_convert_pdf_not_found(self, temp_dir):
        """存在しないPDFファイル"""
        converter = PDFConverter(temp_dir)
        non_existent_pdf = os.path.join(temp_dir, "non_existent.pdf")
        result = converter.convert(non_existent_pdf)
        assert result is None

    def test_convert_image_to_pdf(self, temp_dir, sample_image_file):
        """画像からPDFへの変換"""
        converter = PDFConverter(temp_dir)
        output_path = os.path.join(temp_dir, "output.pdf")
        result = converter.convert(sample_image_file, output_path)
        assert result is not None
        assert os.path.exists(result)

    def test_convert_already_converted(self, temp_dir, sample_image_file):
        """既に変換済みの場合はスキップ"""
        converter = PDFConverter(temp_dir)
        # 最初の変換
        output_path = os.path.join(temp_dir, "sample.png.pdf")
        result1 = converter.convert(sample_image_file)
        assert result1 is not None

        # 2回目の変換（スキップされる）
        result2 = converter.convert(sample_image_file)
        assert result2 == result1

    def test_office_extensions(self):
        """サポートされるOffice拡張子"""
        assert '.doc' in PDFConverter.OFFICE_EXTENSIONS
        assert '.docx' in PDFConverter.OFFICE_EXTENSIONS
        assert '.xls' in PDFConverter.OFFICE_EXTENSIONS
        assert '.xlsx' in PDFConverter.OFFICE_EXTENSIONS
        assert '.ppt' in PDFConverter.OFFICE_EXTENSIONS
        assert '.pptx' in PDFConverter.OFFICE_EXTENSIONS
        assert '.rtf' in PDFConverter.OFFICE_EXTENSIONS

    def test_image_extensions(self):
        """サポートされる画像拡張子"""
        assert '.jpg' in PDFConverter.IMAGE_EXTENSIONS
        assert '.jpeg' in PDFConverter.IMAGE_EXTENSIONS
        assert '.png' in PDFConverter.IMAGE_EXTENSIONS
        assert '.bmp' in PDFConverter.IMAGE_EXTENSIONS
        assert '.tiff' in PDFConverter.IMAGE_EXTENSIONS

    def test_ichitaro_extensions(self):
        """サポートされる一太郎拡張子"""
        assert '.jtd' in PDFConverter.ICHITARO_EXTENSIONS

    def test_convert_rgba_image(self, temp_dir):
        """RGBA画像の変換（RGB変換を含む）"""
        try:
            from PIL import Image
            # RGBA画像を作成
            rgba_path = os.path.join(temp_dir, "rgba_image.png")
            img = Image.new('RGBA', (100, 100), color=(255, 0, 0, 128))
            img.save(rgba_path)
            img.close()

            converter = PDFConverter(temp_dir)
            result = converter.convert(rgba_path)
            assert result is not None
            assert os.path.exists(result)
        except ImportError:
            pytest.skip("PIL/Pillow not installed")


class TestPDFConverterEdgeCases:
    """PDFConverterのエッジケーステスト"""

    def test_empty_temp_dir(self, temp_dir):
        """空の一時ディレクトリ"""
        converter = PDFConverter(temp_dir)
        assert converter.temp_dir == temp_dir

    def test_ichitaro_settings_merge(self, temp_dir):
        """一太郎設定のマージが正しく動作する"""
        partial_settings = {'open_wait_seconds': 15}
        converter = PDFConverter(temp_dir, partial_settings)

        # カスタム値が適用される
        assert converter.ichitaro_settings['open_wait_seconds'] == 15
        # 他のデフォルト値は維持される
        assert converter.ichitaro_settings['dialog_wait_seconds'] == 3
        assert converter.ichitaro_settings['action_wait_seconds'] == 2
