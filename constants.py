"""
定数モジュール

アプリケーション全体で使用する定数を定義
"""
from enum import IntEnum


class WordFormat(IntEnum):
    """Word FileFormat定数"""
    PDF = 17  # wdFormatPDF


class ExcelFormat(IntEnum):
    """Excel FileFormat定数"""
    PDF = 0  # xlTypePDF


class PowerPointFormat(IntEnum):
    """PowerPoint FileFormat定数"""
    PDF = 32  # ppSaveAsPDF


class ExcelLookIn(IntEnum):
    """Excel検索対象の定数"""
    VALUES = -4163  # xlValues


class ExcelLookAt(IntEnum):
    """Excel検索方法の定数"""
    PART = 2   # xlPart (部分一致)
    WHOLE = 1  # xlWhole (完全一致)


class ExcelSortOrder(IntEnum):
    """Excelソート順の定数"""
    ASCENDING = 1   # xlAscending (昇順)
    DESCENDING = 2  # xlDescending (降順)


class ExcelSortHeader(IntEnum):
    """Excelソートヘッダーの定数"""
    YES = 1  # xlYes (ヘッダー行あり)
    NO = 0   # xlNo (ヘッダー行なし)


class ExcelTransferConstants:
    """
    Excel自動転記処理の定数

    注意: このクラスの定数は後方互換性のために残されています。
    新しい実装では config.json の excel_transfer セクションを使用してください。
    """

    # 行事キーワード（イベント判定用）
    EVENT_KEYWORDS = ["儀式", "文化", "保健", "遠足", "勤労", "その", "児童"]

    # 欠時キーワード
    ABSENT_KEYWORD = "欠時"

    # 処理範囲（ループ1：D8～D50）
    LOOP1_START_ROW = 8
    LOOP1_END_ROW = 51  # Python range用（50まで処理）
    LOOP1_SEARCH_COL = "D"
    LOOP1_SORT_RANGE = "A8:P50"
    LOOP1_SORT_KEY = "B8"

    # 処理範囲（ループ2：C55～C62）
    LOOP2_START_ROW = 55
    LOOP2_END_ROW = 63  # Python range用（62まで処理）
    LOOP2_SEARCH_COL = "C"
    LOOP2_SORT_RANGE = "A55:P62"
    LOOP2_SORT_KEY = "B55"

    # 処理範囲（ループ3：C67～C96）
    LOOP3_START_ROW = 67
    LOOP3_END_ROW = 97  # Python range用（96まで処理）
    LOOP3_SEARCH_COL = "C"
    LOOP3_SORT_RANGE = "A67:P96"
    LOOP3_SORT_KEY = "B67"

    # 参照ファイルの列範囲
    REF_DATA_START_COL = "E"  # E列から開始
    REF_DATA_END_COL = "AN"   # AN列まで（36列＝6学年×6校時）
    REF_SEARCH_COL = 3        # C列（行事名検索用）
    REF_DATE_COL = 1          # A列（日付取得用）

    # 学年別のデータ構造
    GRADES_COUNT = 6          # 学年数
    PERIODS_PER_GRADE = 6     # 1学年あたりの校時数
    TOTAL_COLUMNS = 36        # 総列数（6学年×6校時）

    # 学年別マッピング（反映ファイル）
    # {学年: (行事時数の列, 欠時数の列)}
    GRADE_COLUMN_MAPPING = {
        1: ("E", "F"),
        2: ("G", "H"),
        3: ("I", "J"),
        4: ("K", "L"),
        5: ("M", "N"),
        6: ("O", "P")
    }

    # 反映ファイルの列
    TARGET_DATE_COL = "A"           # 日付列
    TARGET_PROCESS_DATE_COL = "B"   # 処理日列
    TARGET_FILTER_COL = "C"         # フィルターキーワード列（D8～D50用）


class AppConstants:
    """アプリケーション定数"""

    # バージョン情報
    VERSION = "3.4.0"
    APP_NAME = "教育計画PDFマージシステム"

    # デフォルトタイムアウト（秒）
    DEFAULT_TIMEOUT_SECONDS = 30

    # 一太郎変換のデフォルト設定
    ICHITARO_DEFAULTS = {
        'ichitaro_ready_timeout': 30,  # 一時ファイル検出の最大待機時間（秒）
        'max_retries': 3,              # 最大リトライ回数
        'save_wait_seconds': 20        # PDF保存待機時間（秒）
    }

    # 一時ファイルのデフォルト保持期間（時間）
    TEMP_FILE_MAX_AGE_HOURS = 24

    # GUIログハンドラーで使用するロガー名リスト
    GUI_LOGGER_NAMES = [
        'pdf_converter',
        'converters.office_converter',
        'converters.image_converter',
        'converters.ichitaro_converter',
        'pdf_processor',
        'document_collector',
        '__main__'
    ]


class PDFConstants:
    """PDF処理に関連する定数"""

    # ページ構成
    COVER_PAGE_COUNT = 1       # 表紙ページ数
    TOC_PAGE_COUNT = 1         # 目次ページ数
    CONTENT_START_PAGE = 3     # コンテンツ開始ページ（表紙 + 目次 + 1）

    # ページ番号の表示設定
    PAGE_NUMBER_X_OFFSET = 10          # 中央からの左オフセット（ポイント）
    PAGE_NUMBER_BOTTOM_MARGIN = 30     # 下端からのマージン（ポイント）
    PAGE_NUMBER_FONT_SIZE = 12         # フォントサイズ（ポイント）
    PAGE_NUMBER_FONT_NAME = "helv"     # フォント名

    # ページ数のデフォルト値
    DEFAULT_PAGE_COUNT = 1             # 取得失敗時のフォールバック

    # 目次見出しレベル
    HEADING_LEVEL_MAIN = 1             # 大見出し（メインディレクトリ）
    HEADING_LEVEL_SUB = 2              # 小見出し（サブフォルダ/ファイル）

    # 目次レイアウト定数
    TOC_TITLE_COL_WIDTH_RATIO = 0.8    # タイトル列の幅（ドキュメント幅に対する比率）
    TOC_PAGE_COL_WIDTH_RATIO = 0.2     # ページ番号列の幅（ドキュメント幅に対する比率）
    TOC_FRAME_ID = 'toc_frame'         # 目次フレームのID

    # ファイル名接尾辞
    TEMP_FILE_SUFFIX = ".tmp"          # 一時ファイルの拡張子

    # Ghostscript圧縮設定
    GS_COMPATIBILITY_LEVEL = "1.4"     # PDF互換性レベル
    GS_PDF_SETTINGS = "/ebook"         # 品質設定
    GS_TIMEOUT_SECONDS = 30            # タイムアウト（秒）


class IchitaroWaitTimes:
    """
    一太郎変換の待機時間定数（環境に応じて config.json で調整可能）

    低スペックPCの場合は config.json の ichitaro セクションで値を増やすことを推奨
    """
    # 起動・接続
    STARTUP_WAIT = 3.0      # 一太郎起動待機時間（秒）

    # 印刷ダイアログ操作
    CTRL_P_WAIT = 3.0       # Ctrl+P後の待機時間（秒）
    PRINTER_SELECT_WAIT = 0.5  # プリンター選択後の待機時間（秒）
    CTRL_A_WAIT = 0.5       # Ctrl+A後の待機時間（秒）
    ENTER_INTERVAL = 0.8    # Enter連打の間隔（秒）

    # 保存ダイアログ操作
    DIALOG_TIMEOUT = 30     # 保存ダイアログ検出のタイムアウト（秒）
    DIALOG_POLL_INTERVAL = 0.3  # ダイアログ検出のポーリング間隔（秒）
    DIALOG_MIN_WAIT = 2.0   # ダイアログ検出開始前の最低待機時間（秒）
    KEYBOARD_PREP_WAIT = 0.3  # キーボード入力準備の待機時間（秒）
    FILE_INPUT_WAIT = 0.5   # ファイルパス入力後の待機時間（秒）

    # プロセス終了
    PRINT_COMPLETE_WAIT = 2.0  # 印刷処理完了待機時間（秒）
    WINDOW_CLOSE_WAIT = 0.5    # ウィンドウクローズ後の待機時間（秒）
    CLEANUP_TIMEOUT = 1     # クリーンアップの接続タイムアウト（秒）
    CLEANUP_WAIT = 0.5      # クリーンアップ後の待機時間（秒）

    # リトライ設定
    MAX_ATTEMPTS = 3        # 一太郎変換の最大試行回数
    RETRY_DELAY = 2.0       # 再試行前の待機時間（秒）


class PathConstants:
    """パス処理関連の定数"""

    # ファイル名のデフォルト値
    DEFAULT_FILENAME = 'file'


class PDFConversionConstants:
    """PDF変換処理の定数"""

    # 動的ファイル待機の間隔設定（秒）
    FILE_WAIT_INTERVAL_FAST = 0.1        # 最初の10回（高速チェック）
    FILE_WAIT_INTERVAL_MEDIUM = 0.5      # 次の20回（中速チェック）
    FILE_WAIT_INTERVAL_SLOW = 1.0        # 残り（低速チェック）
    FILE_WAIT_FAST_COUNT = 10            # 高速チェック回数
    FILE_WAIT_MEDIUM_COUNT = 20          # 中速チェック回数
    FILE_STABILITY_THRESHOLD = 3         # ファイルサイズ安定判定の閾値（回数）
    FILE_WAIT_LOG_INTERVAL = 5           # ファイル待機のログ出力間隔（秒）

    # キャンセルチェック間隔（秒）
    CANCEL_CHECK_INTERVAL = 0.5

    # プリンター選択リトライ
    PRINTER_SELECT_MAX_RETRIES = 3       # プリンター選択の最大リトライ回数
    PRINTER_SELECT_RETRY_DELAY = 1.0     # プリンター選択のリトライ遅延（秒）

    # ログメッセージ
    LOG_MARK_SUCCESS = "✓"
    LOG_MARK_FAILURE = "✗"
    LOG_SEPARATOR_MAJOR = "=" * 60

    # ファイル名のデフォルト値
    DEFAULT_SEPARATOR_NAME = 'separator'
