"""
PDF統合スクリプト

指定されたディレクトリ内の各種ファイル（Office、画像、一太郎、PDF）を
PDFに変換し、目次とページ番号を付けて1つのPDFファイルにまとめます。

使用方法:
    python convert_and_merge.py [--type education|event]

    --type education: 教育計画（デフォルト、サブフォルダに区切りページを作成）
    --type event: 行事計画（サブフォルダには区切りページを作成しない）
"""
import argparse
from config_loader import ConfigLoader
from pdf_converter import PDFConverter
from pdf_processor import PDFProcessor
from document_collector import DocumentCollector, PDFMergeOrchestrator


def main(plan_type: str = "education") -> None:
    """
    メイン処理

    Args:
        plan_type: 計画種別 ("education" または "event")
    """
    # 設定ファイルの読み込み
    config = ConfigLoader()

    # 各モジュールの初期化
    temp_dir = config.get_temp_dir()
    converter = PDFConverter(temp_dir)
    processor = PDFProcessor(config)
    template_path = config.get_template_path()
    collector = DocumentCollector(converter, processor, template_path)

    # オーケストレーターの初期化
    orchestrator = PDFMergeOrchestrator(config, converter, processor, collector)

    # パスの取得（計画種別に応じて切り替え）
    if plan_type == "event":
        target_directory = config.get_event_plan_path()
        create_separator_for_subfolder = False
    else:
        target_directory = config.get_education_plan_path()
        create_separator_for_subfolder = True

    output_pdf_file = config.get('output', 'merged_pdf')

    # PDF統合処理の実行
    orchestrator.create_merged_pdf(
        target_dir=target_directory,
        output_pdf=output_pdf_file,
        create_separator_for_subfolder=create_separator_for_subfolder
    )


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description="PDF統合スクリプト - 各種ファイルをPDFに変換して1つにまとめます"
    )
    parser.add_argument(
        '--type', '-t',
        choices=['education', 'event'],
        default='education',
        help='計画種別: education=教育計画（デフォルト）, event=行事計画'
    )
    args = parser.parse_args()
    main(plan_type=args.type)
