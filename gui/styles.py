"""
UI定数モジュール

色、フォント、ウィンドウサイズなどのUI定数を定義
"""

# カラーパレット
COLORS = {
    'primary': '#4CAF50',
    'primary_hover': '#45a049',
    'secondary': '#2196F3',
    'secondary_hover': '#1976D2',
    'warning': '#FF9800',
    'warning_hover': '#F57C00',
    'error': '#f44336',
    'error_hover': '#d32f2f',
    'success': '#4CAF50',
    'info': '#2196F3',
    'background': '#f5f5f5',
    'surface': '#ffffff',
    'text_primary': '#212121',
    'text_secondary': '#757575',
    'border': '#e0e0e0',
}

# フォント設定
FONTS = {
    'default': ('メイリオ', 10),
    'title': ('メイリオ', 11, 'bold'),
    'small': ('メイリオ', 9),
    'mono': ('Consolas', 10),
}

# ウィンドウ設定
# 注: タイトルはconstants.AppConstants.APP_NAMEを使用
WINDOW = {
    'geometry': '950x780',
    'min_width': 800,
    'min_height': 650,
}

# ボタンスタイル
BUTTON_STYLES = {
    'primary': {
        'bg': COLORS['primary'],
        'fg': 'white',
        'activebackground': COLORS['primary_hover'],
    },
    'secondary': {
        'bg': COLORS['secondary'],
        'fg': 'white',
        'activebackground': COLORS['secondary_hover'],
    },
    'warning': {
        'bg': COLORS['warning'],
        'fg': 'white',
        'activebackground': COLORS['warning_hover'],
    },
    'error': {
        'bg': COLORS['error'],
        'fg': 'white',
        'activebackground': COLORS['error_hover'],
    },
}

# パディング・マージン
PADDING = {
    'small': 5,
    'medium': 10,
    'large': 15,
    'xlarge': 20,
}
