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

    # Sort Order定数
    XL_ASCENDING = 1  # xlAscending (昇順)
    XL_DESCENDING = 2  # xlDescending (降順)

    # Sort Header定数
    XL_YES = 1  # xlYes (ヘッダー行あり)
    XL_NO = 0   # xlNo (ヘッダー行なし)


class ExcelTransferConstants:
    """Excel自動転記処理の定数"""

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
    VERSION = "3.2.4"  # 一太郎変換改善版（警告ダイアログ＆リトライ機能）
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


class PDFConversionConstants:
    """PDF変換処理の定数（マジックナンバー削減）"""

    # 一太郎変換の待機時間（秒）
    ICHITARO_STARTUP_WAIT = 3.0          # 一太郎起動待機時間
    ICHITARO_PRINT_DIALOG_WAIT = 3.0     # 印刷ダイアログ待機時間（低スペックPC対応）
    ICHITARO_SAVE_DIALOG_WAIT = 3.0      # 保存ダイアログ表示待機時間
    ICHITARO_CTRL_P_WAIT = 3.0           # Ctrl+P送信後の待機時間（低スペックPC対応）
    ICHITARO_ENTER_INTERVAL = 0.8        # Enter連打の間隔
    ICHITARO_PRINT_COMPLETE_WAIT = 2.0   # 印刷処理完了待機時間
    ICHITARO_WINDOW_CLOSE_WAIT = 1.0     # ウィンドウクローズ待機時間
    ICHITARO_PRINTER_SELECT_WAIT = 0.5   # プリンタ選択後の待機時間
    ICHITARO_FILE_INPUT_WAIT = 0.5       # ファイル名入力後の待機時間
    ICHITARO_CTRL_A_WAIT = 0.3           # Ctrl+A後の待機時間
    ICHITARO_RETRY_DELAY = 2.0           # 再試行前の待機時間
    ICHITARO_DIALOG_MIN_WAIT = 3.0       # 保存ダイアログ検出の最低待機時間
    ICHITARO_DIALOG_TIMEOUT = 30         # 保存ダイアログ検出のタイムアウト
    ICHITARO_DIALOG_POLL_INTERVAL = 0.5  # 保存ダイアログ検出のポーリング間隔
    ICHITARO_KEYBOARD_PREP_WAIT = 0.5    # キーボード入力準備の待機時間
    ICHITARO_CLEANUP_TIMEOUT = 1         # クリーンアップの接続タイムアウト
    ICHITARO_CLEANUP_WAIT = 0.5          # クリーンアップ後の待機時間

    # プリンター選択のリトライ設定
    PRINTER_SELECT_MAX_RETRIES = 3       # プリンター選択の最大リトライ回数
    PRINTER_SELECT_RETRY_DELAY = 1.0     # プリンター選択のリトライ遅延

    # 一太郎変換の試行回数
    ICHITARO_MAX_ATTEMPTS = 3            # 一太郎変換の最大試行回数（初回を含む）

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

    # ファイル名サニタイズのデフォルト名
    DEFAULT_SEPARATOR_NAME = 'separator'

    # ログメッセージの装飾記号（統一）
    LOG_SEPARATOR_MAJOR = "=" * 60       # 主要セクションの区切り線
    LOG_SEPARATOR_MINOR = "-" * 60       # 副セクションの区切り線
    LOG_MARK_SUCCESS = "✓"               # 成功マーク
    LOG_MARK_FAILURE = "✗"               # 失敗マーク
