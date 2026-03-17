"""
UI定数モジュール

色、フォント、ウィンドウサイズ、ボタンスタイルを定義
"""

# カラーパレット
COLORS = {
    # Primary
    'primary': '#1976D2',
    'primary_hover': '#1565C0',
    'primary_light': '#BBDEFB',

    # Semantic
    'success': '#388E3C',
    'success_hover': '#2E7D32',
    'error': '#D32F2F',
    'error_hover': '#C62828',

    # Surface
    'surface': '#FFFFFF',
    'surface_dim': '#f0f0f0',
    'surface_light': '#F0F8FF',

    # Text
    'text_primary': '#212121',
    'text_secondary': '#757575',

    # Validation
    'valid': '#4CAF50',
    'invalid': '#F44336',

    # Warning
    'warning_mild': '#FF9800',

    # Tooltip
    'tooltip_bg': '#FFFFE0',
    'tooltip_fg': '#000000',
}

# フォント設定
FONTS = {
    'heading': ('メイリオ', 12, 'bold'),
    'subheading': ('メイリオ', 11, 'bold'),
    'default': ('メイリオ', 10),
    'default_bold': ('メイリオ', 10, 'bold'),
    'small': ('メイリオ', 9),
    'small_bold': ('メイリオ', 9, 'bold'),
    'tiny': ('メイリオ', 8),
    'status_dot': ('メイリオ', 12),
    'dialog_title': ('メイリオ', 16, 'bold'),
    'dialog_heading': ('メイリオ', 14, 'bold'),
}

# ウィンドウ設定
WINDOW = {
    'geometry': '1024x768',
    'min_width': 800,
    'min_height': 600,
}

# ボタンスタイル
BUTTON_STYLES = {
    'primary': {
        'bg': COLORS['primary'],
        'fg': 'white',
        'activebackground': COLORS['primary_hover'],
    },
    'success': {
        'bg': COLORS['success'],
        'fg': 'white',
        'activebackground': COLORS['success_hover'],
    },
}
