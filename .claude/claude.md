# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## プロジェクト概要

教育計画・行事計画のドキュメント（Word、Excel、PowerPoint、一太郎、画像、PDF）を目次・ブックマーク付き単一PDFに統合するWindowsデスクトップアプリ。Excel自動転記機能も搭載。

**制約**: Windows専用（Win32 COM、pywinauto使用）。Office必須。一太郎変換はUI自動化。

## 開発コマンド

```bash
# セットアップ
python -m venv venv && venv\Scripts\activate
pip install -r requirements-dev.txt
pre-commit install

# 実行
python run_app.py

# テスト
pytest                    # 全テスト（325件）
pytest -m unit            # ユニットテストのみ（285件）
pytest tests/test_pdf_converter.py::TestConvert::test_word  # 単一テスト

# 品質チェック（コミット前に必ず実行）
pre-commit run --all-files    # ruff + mypy + bandit を一括実行

# ビルド
pyinstaller build_installer.spec --clean --noconfirm
# インストーラー: Inno Setup 6が必要
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer\setup.iss
```

## コミットルール

**`--no-verify` は絶対に使わない。** pre-commitフックを必ず通す。

```text
種別: 簡潔な説明
Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>
```

種別: `Feature`, `Fix`, `Refactor`, `Docs`, `Test`, `Chore`, `Build`

## アーキテクチャ

```text
gui/                    # tkinter GUI（BaseTab継承、styles.pyで定数管理）
  ├─ tabs/              # PDFTab, ExcelTab, SettingsTab
  ├─ event_names_editor.py
  └─ setup_wizard.py
core/                   # ビジネスロジック
  ├─ pdf_merge_orchestrator.py  # 7ステップのPDF統合フロー制御
  ├─ document_collector.py      # ディレクトリ走査・目次生成
  ├─ pdf_converter.py           # 変換ファサード（converters/に委譲）
  ├─ pdf_processor.py           # PDF操作（マージ、TOC、ブックマーク）
  └─ update_excel_files.py      # Excel転記（COM、キャッシュ、あいまい検索）
converters/             # 各形式→PDF変換
  ├─ office_converter.py        # Word/Excel/PPT（COM）
  ├─ ichitaro_converter.py      # 一太郎（pywinauto UI自動化）
  └─ image_converter.py         # 画像（Pillow）
infrastructure/         # 設定・ユーティリティ
  ├─ config_loader.py           # 2層設定（config.json + user_config.json）
  ├─ ghostscript.py             # GS検出・検証・ダウンロード案内
  └─ path_validator.py          # パス検証・サニタイズ・トラバーサル対策
shared/                 # 全層共通
  ├─ constants.py               # 定数クラス群
  └─ exceptions.py              # PDFMergeError階層（9種類、例外チェーン対応）
```

### 主要な処理フロー

`PDFMergeOrchestrator.create_merged_pdf()` が7ステップで制御:

1. ドキュメント収集・変換 → 2. 一時マージ → 3. 目次PDF生成 → 4. 表紙分割 → 5. 最終マージ → 6. ページ番号・ブックマーク → 7. Ghostscript圧縮

### 一太郎変換の仕組み（重要）

`IchitaroConverter` はpywinautoでUI自動化:

1. **デフォルトプリンターを「Microsoft Print to PDF」に変更**（一太郎起動前に実行が必須）
2. 一太郎でファイルを開く
3. 予期しないダイアログを自動閉じ（`_dismiss_unexpected_dialogs`: いいえ > OK > Enter の優先順）
4. Ctrl+P → Enter（プリンター選択済みなのでそのまま印刷）
5. 印刷後ダイアログ処理（ページ番号確認等）
6. 保存ダイアログでクリップボード経由でパス入力（`send_keys`は日本語特殊文字が消えるため不可）
7. 一太郎終了後にデフォルトプリンターを復元

**注意**: 変換失敗時はスキップせずPDFConversionErrorで処理全体を停止する。

### 出力ファイル名のピリオド

`pdf_converter.py` で出力ファイル名のピリオドをアンダースコアに置換する（`08_4.学校.jtd` → `08_4_学校_xxxx.pdf`）。保存ダイアログが拡張子と誤認するため。

## コーディング標準

- **型ヒント**: すべての関数に必須。`mypy.ini`で厳密度制御（core/shared/infrastructure/convertersは厳密、guiは緩和）
- **例外**: `shared.exceptions` のカスタム例外を使用、`original_error`で例外チェーン
- **パス検証**: ユーザー入力パスは`PathValidator`経由。GUI層では`BaseTab.validate_path()`を使用
- **COM初期化**: `run_app.py`で`sys.coinit_flags = 2`（STA）が必須（tkinter.filedialogとの競合回避）
- **リソース管理**: `fitz.open()`等はコンテキストマネージャー使用。COM は finally で cleanup

## 設定管理

2層構造: `config.json`（デフォルト） + `%LOCALAPPDATA%\PDFMergeSystem\user_config.json`（ユーザー設定）。`ConfigLoader`がディープマージ。

Excel転記の行事名は `ConfigLoader.get_event_names()` / `save_event_names()` で管理。転記実行時に `populate_event_names()` でターゲットExcelに自動設定。

## テスト

- ユニットテストに `@pytest.mark.unit`、統合テストに `@pytest.mark.integration` マーカー付与済み
- COM操作・pywinauto・subprocessはモック使用
- `tests/conftest.py` にfixture定義（temp_dir, sample_config_data, sample_pdf_file等）

## CI/CD

`.github/workflows/ci.yml` で3ジョブ（品質チェック→テスト→ビルド検証）。品質チェック失敗でビルドをブロック。

## GUI

- `gui/styles.py`: COLORS, FONTS, BUTTON_STYLES を一元管理。色名リテラル（`"gray"`, `"white"`等）はハードコードOK（可読性優先）
- `BaseTab`: ログフレーム、折りたたみセクション、パス検証ヘルパーを提供
- ログウィジェットは `fill="both", expand=True` でリサイズ追従
