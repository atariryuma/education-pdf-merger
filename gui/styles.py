"""
UI定数モジュール - Material Design 3準拠

色、フォント、ウィンドウサイズなどのUI定数を定義
"""

# カラーパレット - Material Design 3
COLORS = {
    # Primary colors
    'primary': '#1976D2',           # Material Blue 700
    'primary_hover': '#1565C0',     # Material Blue 800
    'primary_light': '#BBDEFB',     # Material Blue 100

    # Secondary colors
    'secondary': '#0288D1',         # Material Light Blue 700
    'secondary_hover': '#0277BD',   # Material Light Blue 800
    'secondary_light': '#B3E5FC',   # Material Light Blue 100

    # Accent colors
    'accent': '#FF6F00',            # Material Orange 900
    'accent_hover': '#E65100',      # Material Orange 900

    # Semantic colors
    'success': '#388E3C',           # Material Green 700
    'success_hover': '#2E7D32',     # Material Green 800
    'warning': '#F57C00',           # Material Orange 700
    'warning_hover': '#EF6C00',     # Material Orange 800
    'error': '#D32F2F',             # Material Red 700
    'error_hover': '#C62828',       # Material Red 800
    'info': '#0288D1',              # Material Light Blue 700

    # Background & Surface
    'background': '#FAFAFA',        # Material Grey 50
    'surface': '#FFFFFF',           # White
    'surface_variant': '#F5F5F5',   # Material Grey 100

    # Text colors
    'text_primary': '#212121',      # Material Grey 900
    'text_secondary': '#757575',    # Material Grey 600
    'text_disabled': '#BDBDBD',     # Material Grey 400
    'text_hint': '#9E9E9E',         # Material Grey 500

    # Border & Divider
    'border': '#E0E0E0',            # Material Grey 300
    'divider': '#BDBDBD',           # Material Grey 400

    # Overlay
    'overlay': 'rgba(0,0,0,0.5)',
    'scrim': 'rgba(0,0,0,0.32)',
}

# フォント設定 - モダンなシステムフォント
FONTS = {
    'default': ('メイリオ', 10),
    'title': ('メイリオ', 12, 'bold'),
    'subtitle': ('メイリオ', 11),
    'small': ('メイリオ', 9),
    'tiny': ('メイリオ', 8),
    'mono': ('Consolas', 10),
    'button': ('メイリオ', 10, 'bold'),
}

# ウィンドウ設定
# 注: タイトルはconstants.AppConstants.APP_NAMEを使用
WINDOW = {
    'geometry': '950x750',          # UI内容に最適化されたサイズ
    'min_width': 850,
    'min_height': 650,
}

# ボタンスタイル - Material Design 3
BUTTON_STYLES = {
    'primary': {
        'bg': COLORS['primary'],
        'fg': 'white',
        'activebackground': COLORS['primary_hover'],
        'relief': 'flat',
        'borderwidth': 0,
        'highlightthickness': 0,
    },
    'secondary': {
        'bg': COLORS['secondary'],
        'fg': 'white',
        'activebackground': COLORS['secondary_hover'],
        'relief': 'flat',
        'borderwidth': 0,
        'highlightthickness': 0,
    },
    'accent': {
        'bg': COLORS['accent'],
        'fg': 'white',
        'activebackground': COLORS['accent_hover'],
        'relief': 'flat',
        'borderwidth': 0,
        'highlightthickness': 0,
    },
    'success': {
        'bg': COLORS['success'],
        'fg': 'white',
        'activebackground': COLORS['success_hover'],
        'relief': 'flat',
        'borderwidth': 0,
        'highlightthickness': 0,
    },
    'warning': {
        'bg': COLORS['warning'],
        'fg': 'white',
        'activebackground': COLORS['warning_hover'],
        'relief': 'flat',
        'borderwidth': 0,
        'highlightthickness': 0,
    },
    'error': {
        'bg': COLORS['error'],
        'fg': 'white',
        'activebackground': COLORS['error_hover'],
        'relief': 'flat',
        'borderwidth': 0,
        'highlightthickness': 0,
    },
    'outlined': {
        'bg': COLORS['surface'],
        'fg': COLORS['primary'],
        'activebackground': COLORS['primary_light'],
        'relief': 'solid',
        'borderwidth': 2,
        'highlightthickness': 0,
    },
}

# パディング・マージン - 8dpグリッドシステム
PADDING = {
    'xs': 4,      # Extra Small
    'small': 8,   # Small
    'medium': 16, # Medium
    'large': 24,  # Large
    'xlarge': 32, # Extra Large
    'xxlarge': 48,# 2X Large
}

# 角丸設定
RADIUS = {
    'small': 4,
    'medium': 8,
    'large': 12,
    'xlarge': 16,
}

# エレベーション（影）- Material Design
ELEVATION = {
    'none': {},
    'low': {'borderwidth': 1, 'relief': 'solid'},
    'medium': {'borderwidth': 2, 'relief': 'raised'},
    'high': {'borderwidth': 3, 'relief': 'raised'},
}
