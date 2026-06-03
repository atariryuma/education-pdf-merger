# 教育計画PDFマージシステム

[![Python](https://img.shields.io/badge/python-3.10%2B-blue)](https://www.python.org/downloads/)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)](https://www.microsoft.com/windows)

教育計画・行事計画関連のドキュメント（Word／Excel／PowerPoint／一太郎／画像／PDF）を、目次・しおり付きの単一PDFに統合する Windows デスクトップアプリ。Excel自動転記機能も搭載。

## 主な機能

### PDF統合

- **複数形式対応**: Word (.doc/.docx)、Excel (.xls/.xlsx)、PowerPoint (.ppt/.pptx)、一太郎 (.jtd)、画像 (.jpg/.png 等)、PDF
- **目次自動生成**: ディレクトリ構造から階層目次を自動作成
- **しおり付加**: クリック可能なPDFアウトラインを自動設定
- **ページ番号**: 表紙を除いて自動付与（複数ページ表紙にも対応）
- **Ghostscript圧縮**: PDFサイズを最大70%削減
- **区切りページ**: フォルダごとに見出しページを自動挿入

### Excel処理

- **データ自動転記**: 年間行事計画から様式4へのデータ転記
- **時数自動計算**: 学年別の行事時数を自動カウント
- **行事名管理**: 設定画面で行事名を編集・保存

## インストール

### 前提条件

- **Python 3.10以上** — [ダウンロード](https://www.python.org/downloads/)
- **Microsoft Office** — Word／Excel／PowerPoint が必要
- **Ghostscript** — [ダウンロード](https://ghostscript.com/releases/gsdnld.html)
- **一太郎**（任意） — `.jtd` ファイルを扱う場合のみ

### 方法1: インストーラー（推奨）

1. `PDFMergeSystem_Setup_x.x.x.exe` をダウンロード
2. インストーラーを実行
3. スタートメニューまたはデスクトップアイコンから起動

### 方法2: ソースから実行（開発者向け）

```bat
python -m venv venv
venv\Scripts\activate
pip install -r requirements-dev.txt
pre-commit install
python run_app.py
```

## 使い方

GUIアプリを起動:

```bat
python run_app.py
```

### PDF統合タブ

1. 入力ディレクトリを選択（変換したいファイルがあるフォルダ）
2. 出力PDFファイル名を指定
3. 「PDF統合を実行」をクリック

### Excel処理タブ

1. 対象のExcelファイルを開いた状態で実行
2. 「Excelデータ更新を実行」をクリック

### 設定タブ

- 年度の更新（西暦を入力すると和暦は自動計算）
- パス設定・行事名の編集・Ghostscript の検出

## 設定ファイル

設定は2層で管理:

- `config.json`（プロジェクト同梱、デフォルト値）
- `%LOCALAPPDATA%\PDFMergeSystem\user_config.json`（ユーザー個別設定）

`ConfigLoader` が2層をディープマージ。GUIの「設定」タブから編集可能。

## プロジェクト構成

```text
education-pdf-merger/
├─ run_app.py                    アプリケーション起動エントリポイント
├─ build.bat                     EXEビルドスクリプト
├─ build_installer.spec          PyInstaller設定
├─ config.json                   デフォルト設定
├─ pyproject.toml                プロジェクトメタデータ（バージョン管理の単一の真実）
├─ requirements.txt              本番依存
├─ requirements-dev.txt          開発依存
│
├─ core/                         ビジネスロジック層
│   ├─ pdf_merge_orchestrator.py    7ステップのPDF統合フロー制御
│   ├─ document_collector.py        ディレクトリ走査・目次生成
│   ├─ pdf_converter.py             変換ファサード（converters/ に委譲）
│   ├─ pdf_processor.py             PDF操作（マージ・TOC・しおり）
│   ├─ update_excel_files.py        Excel転記（COM、キャッシュ、あいまい検索）
│   └─ folder_structure_detector.py 教育/行事プラン自動判定
│
├─ converters/                   各形式 → PDF 変換
│   ├─ office_converter.py          Word/Excel/PPT（COM）
│   ├─ image_converter.py           画像（Pillow）
│   └─ ichitaro_converter.py        一太郎（pywinauto UI自動化）
│
├─ infrastructure/               設定・ユーティリティ
│   ├─ config_loader.py             2層設定の読み込み
│   ├─ config_validator.py          設定の検証
│   ├─ path_validator.py            パス検証・サニタイズ
│   ├─ ghostscript.py               GS検出・検証
│   ├─ logging_config.py            ロギング設定
│   ├─ year_utils.py                和暦・年度変換
│   └─ com_utils.py                 pythoncom 初期化のコンテキストマネージャ
│
├─ shared/                       全層共通
│   ├─ constants.py                 定数クラス群
│   └─ exceptions.py                例外階層（PDFMergeError）
│
├─ gui/                          tkinter GUI
│   ├─ app.py                       メインウィンドウ
│   ├─ styles.py                    色・フォント・スタイル定数
│   ├─ utils.py                     GUI共通ユーティリティ
│   ├─ setup_wizard.py              初回セットアップ
│   ├─ event_names_editor.py        行事名編集ダイアログ
│   └─ tabs/                        BaseTab、PDFTab、ExcelTab、SettingsTab
│
├─ tests/                        335件のテスト（pytest、@pytest.mark.unit）
└─ installer/                    Inno Setup によるインストーラー
```

## アーキテクチャ

### 処理フロー

```text
1. ドキュメント収集 → 各形式をPDFに変換
2. 一時マージ
3. 目次PDF生成
4. 表紙とコンテンツに分割
5. 表紙 + 目次 + コンテンツを最終マージ
6. ページ番号付加・しおり設定
7. Ghostscript圧縮（オプション）
```

`core/pdf_merge_orchestrator.py` の `create_merged_pdf()` が全体を制御。

### レイヤー構造

```text
gui/  →  core/  →  converters/, infrastructure/  →  shared/
```

- `gui/` は `core/` と `infrastructure/` を呼び出すが、`converters/` を直接呼ばない
- `shared/` はどの層からも参照される（定数・例外）

## 開発

### テスト

```bat
pytest                                              # 全テスト（335件）
pytest -m unit                                      # ユニットテストのみ
pytest tests/test_pdf_processor_ops.py -v           # 単一ファイル
```

### 品質チェック

```bat
pre-commit run --all-files                          # ruff + mypy + bandit 一括
```

`--no-verify` は絶対に使用しない。pre-commit フックを必ず通すこと。

### ビルド

```bat
build.bat                                           # EXE生成
cd installer && build_installer.bat                 # インストーラー生成（Inno Setup 6必要）
```

### コーディング規約

- 型ヒント必須（`mypy.ini` で `core/shared/infrastructure/converters` は厳密、`gui` は緩和）
- 例外: `shared.exceptions` のカスタム例外を使い、`original_error` で例外チェーン
- パス検証: `PathValidator` を使用、GUI層では `BaseTab.validate_path()`
- COM操作: バックグラウンドスレッドでは `infrastructure.com_utils.com_apartment(sta=True)` で囲む
- PDF操作: `pypdf>=4.0.0` を使用（PyPDF2は非推奨）

詳細は [.claude/CLAUDE.md](.claude/CLAUDE.md) を参照。

### コミット規約

```text
種別: 簡潔な説明

詳細（任意、複数行可）

Co-Authored-By: ...
```

種別: `Feature`, `Fix`, `Refactor`, `Docs`, `Test`, `Chore`, `Build`

## ライセンス

MIT License — 詳細は [LICENSE](LICENSE) 参照。

## トラブルシューティング

| 症状 | 原因 / 対策 |
| --- | --- |
| `config.json not found` | プロジェクトルートに `config.json` が無い → 同梱版をコピー |
| PDF変換失敗 | Office／一太郎／Ghostscript 未インストール → 必要なソフトをインストール |
| 一太郎変換が止まる | UI自動化のタイミング問題 → `config.json` の `ichitaro` セクションで調整 |
| Excelスクリプトが動かない | 対象Excelが開かれていない／シート名不一致 → ファイル名・シート名を確認 |
| 変換が遅い | Ghostscript圧縮に時間がかかる → 圧縮スキップ、または小さいファイルで検証 |

## 参考資料

- [pypdf ドキュメント](https://pypdf.readthedocs.io/)
- [PyMuPDF (fitz) ドキュメント](https://pymupdf.readthedocs.io/)
- [ReportLab ユーザーガイド](https://www.reportlab.com/docs/reportlab-userguide.pdf)
- [pywin32 ドキュメント](https://github.com/mhammond/pywin32)
- [pywinauto ドキュメント](https://pywinauto.readthedocs.io/)
- [Inno Setup 公式](https://jrsoftware.org/isinfo.php)

## 変更履歴

詳細は [CHANGELOG.md](CHANGELOG.md) を参照。
