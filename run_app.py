"""
教育計画PDFマージシステム - エントリーポイント

PyInstallerでのビルド用メインスクリプト
"""
import sys
import os

# プロジェクトルートをパスに追加
if getattr(sys, 'frozen', False):
    # PyInstallerでビルドされた場合
    application_path = os.path.dirname(sys.executable)
else:
    # 通常の実行
    application_path = os.path.dirname(os.path.abspath(__file__))

sys.path.insert(0, application_path)

# ロギングの設定
import logging
from logging_config import setup_logging
setup_logging(level=logging.DEBUG)

# アプリケーションの起動（遅延ロード）
def main():
    """メインアプリケーションを起動"""
    from gui.app import main as app_main
    app_main()

if __name__ == "__main__":
    main()
