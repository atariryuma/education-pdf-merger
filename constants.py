"""
定数モジュール

アプリケーション全体で使用する定数を定義
"""


class MSOfficeConstants:
    """Microsoft Office関連の定数"""

    # Word FileFormat定数
    WORD_PDF_FORMAT = 17  # wdFormatPDF

    # Excel FileFormat定数
    EXCEL_PDF_FORMAT = 0  # xlTypePDF

    # PowerPoint FileFormat定数
    POWERPOINT_PDF_FORMAT = 32  # ppSaveAsPDF


class ExcelConstants:
    """Excel COM自動化用の定数"""

    # LookIn定数
    XL_VALUES = -4163  # xlValues

    # LookAt定数
    XL_PART = 2  # xlPart (部分一致)
    XL_WHOLE = 1  # xlWhole (完全一致)


class AppConstants:
    """アプリケーション定数"""

    # バージョン情報
    VERSION = "3.2"
    APP_NAME = "教育計画PDFマージシステム"

    # デフォルトタイムアウト（秒）
    DEFAULT_TIMEOUT_SECONDS = 30

    # 一太郎変換のデフォルト設定
    ICHITARO_DEFAULTS = {
        'ichitaro_ready_timeout': 30,  # 一時ファイル検出の最大待機時間（秒）
        'max_retries': 3,              # 最大リトライ回数
        'down_arrow_count': 5,         # プリンタ選択の下矢印回数
        'save_wait_seconds': 20        # PDF保存待機時間（秒）
    }

    # 一時ファイルのデフォルト保持期間（時間）
    TEMP_FILE_MAX_AGE_HOURS = 24
