"""
pytest fixtures

テスト共通の設定とフィクスチャを定義
"""
import os
import json
import tempfile
import pytest


@pytest.fixture
def temp_dir():
    """一時ディレクトリを作成"""
    with tempfile.TemporaryDirectory() as tmpdir:
        yield tmpdir


@pytest.fixture
def sample_config_data():
    """サンプル設定データ"""
    return {
        "year": "令和７年度(2025)",
        "year_short": "R7",
        "base_paths": {
            "google_drive": "C:\\TestDrive",
            "network": "\\\\server\\share",
            "local_temp": "C:\\Temp\\PDFMerge"
        },
        "directories": {
            "education_plan_base": "教育計画",
            "education_plan": "教育計画書",
            "event_plan": "行事計画",
            "guidance_plan": "指導計画"
        },
        "files": {
            "separator_template": "区切りページ.docx",
            "excel_reference": "参照元.xlsx",
            "excel_target": "対象.xlsx",
            "excel_reference_sheet": "Sheet1",
            "excel_target_sheet": "Sheet1"
        },
        "output": {
            "merged_pdf": "統合PDF.pdf"
        },
        "fonts": {
            "mincho": "C:\\Windows\\Fonts\\msmincho.ttc"
        },
        "ghostscript": {
            "executable": "C:\\Program Files\\gs\\bin\\gswin64c.exe"
        },
        "ichitaro": {
            "open_wait_seconds": 5,
            "dialog_wait_seconds": 3,
            "action_wait_seconds": 2,
            "short_wait_seconds": 1,
            "connect_timeout_seconds": 30,
            "max_retries": 3
        }
    }


@pytest.fixture
def config_file(temp_dir, sample_config_data):
    """設定ファイルを作成"""
    config_path = os.path.join(temp_dir, "config.json")
    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(sample_config_data, f, ensure_ascii=False, indent=2)
    return config_path


@pytest.fixture
def sample_pdf_file(temp_dir):
    """サンプルPDFファイルを作成（ダミー）"""
    pdf_path = os.path.join(temp_dir, "sample.pdf")
    # 最小限のPDFヘッダー（テスト用ダミー）
    with open(pdf_path, 'wb') as f:
        f.write(b'%PDF-1.4\n%\xe2\xe3\xcf\xd3\n')
        f.write(b'1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n')
        f.write(b'2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n')
        f.write(b'3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] >>\nendobj\n')
        f.write(b'xref\n0 4\n0000000000 65535 f \n')
        f.write(b'0000000015 00000 n \n0000000068 00000 n \n0000000131 00000 n \n')
        f.write(b'trailer\n<< /Size 4 /Root 1 0 R >>\nstartxref\n214\n%%EOF\n')
    return pdf_path


@pytest.fixture
def sample_image_file(temp_dir):
    """サンプル画像ファイルを作成"""
    try:
        from PIL import Image
        img_path = os.path.join(temp_dir, "sample.png")
        img = Image.new('RGB', (100, 100), color='red')
        img.save(img_path)
        img.close()
        return img_path
    except ImportError:
        pytest.skip("PIL/Pillow not installed")
