# 変更履歴

## [3.2.4] - 2025-12-25

### 🚀 主要機能改善

- **一太郎変換中の警告ダイアログ** (gui/ichitaro_dialog.py)
  - 変換中に常に最前面で警告ダイアログを表示
  - キーボード入力干渉を防ぐためのユーザー通知
  - 非モーダルでGUIイベントループをブロックしない
  - リトライ状況をリアルタイム表示

- **一太郎変換の自動リトライ機能** (pdf_converter.py)
  - 変換失敗時に最大3回まで自動リトライ
  - リトライ前に一太郎プロセスを強制クリーンアップ
  - リトライ状況をログに詳細表示
  - 3回失敗後はスキップして次のファイルへ

### 🔧 技術的改善

- **非モーダルダイアログ設計**
  - `topmost=True` で常に最前面
  - `grab_set()` を使わずGUIイベントループを維持
  - スレッドセーフな更新（thread_safe_call使用）

- **リトライロジックの実装**
  - 既存のステップ1-5構造を保持しつつリトライループ追加
  - `_cleanup_ichitaro_windows()` による冪等なクリーンアップ
  - CancelledErrorを適切に伝播してキャンセル機構と連携

### 📝 変更ファイル

- gui/ichitaro_dialog.py: 新規作成（警告ダイアログクラス）
- pdf_converter.py: dialog_callback追加、_convert_ichitaro()リトライ実装
- gui/tabs/pdf_tab.py: ダイアログ管理コード追加
- constants.py: VERSION 3.2.4

## [3.2.3] - 2025-12-24

### 🚀 主要機能改善

- **区切りページの自動生成** (pdf_processor.py, pdf_converter.py)
  - Separator.docxテンプレートを廃止し、reportlabで完全自動生成
  - 目次生成と同じ方式で統一（シンプルなタイトル中央配置）
  - Word COM依存を削減、環境非依存に
  - テンプレートファイル管理が不要に

### 🔧 技術的改善

- **コード簡素化とパフォーマンス向上**
  - Word COM操作（約60行）を削除
  - reportlab直接生成で高速化
  - 目次と区切りページの実装パターンを統一
  - PDFProcessor.create_separator_pdf()メソッド追加

### 📝 変更ファイル

- pdf_processor.py: create_separator_pdf()メソッド追加
- pdf_converter.py: create_separator_page()をreportlab版に置き換え
- document_collector.py: template_pathパラメータを削除
- gui/tabs/pdf_tab.py: テンプレートパス取得を削除
- config_loader.py: get_template_path()メソッドを削除

## [3.2.2] - 2025-12-24

### 🚀 主要機能改善

- **一太郎PDF変換の自動プリンター選択（ベストプラクティス実装）** (pdf_converter.py)
  - **pywinauto `select()` メソッドでユーザー操作を完全シミュレート**
  - Microsoft Print to PDFをプリンター名で直接選択（環境非依存）
  - プリンターの並び順に依存しない堅牢な実装
  - 低スペックPC対応：リトライ機構（最大3回）と待機時間延長
  - コード量を60行以上→18行に削減、保守性が大幅向上
  - **下矢印キー方式を完全廃止**（設定項目から削除）
  - 様々なスペックのPCで安定動作

### ✨ UI/UX改善

- **設定タブのスクロール修正** (gui/tabs/settings_tab.py)
  - マウスホイールスクロールが正常に動作するよう修正
  - 再帰的バインドですべての子ウィジェット（Entry、Button等）に対応
  - Entry/Buttonの上でもスクロール可能に
  - `return "break"`でイベント伝播を制御

- **PathValidatorの導入** (path_validator.py, gui/tabs/settings_tab.py)
  - pathlibベースのモダンなパス検証
  - セキュリティ強化（ディレクトリトラバーサル対策）
  - Google DriveとNetwork Driveに「開く」ボタン追加
  - Excelファイル参照ボタンの削除（ファイル名のみの入力フィールド）

- **詳細設定のデフォルト展開** (gui/tabs/settings_tab.py)
  - 一太郎設定がすぐに見えるよう改善
  - プリンター自動選択の説明を表示

### 🔧 技術的改善

- **pywinautoベストプラクティスの適用** (pdf_converter.py)
  - 低レベルWin32 APIから高レベルAPIへ移行
  - `select()` メソッドで必要な通知が自動送信される
  - CB_SETCURSEL/CBN_SELENDOKの手動処理を廃止
  - シンプルで保守しやすいコードに改善

### 🛠️ インストーラー改善

- **プロセス管理の強化** (installer/setup.iss)
  - インストール前にアプリケーションの実行チェック
  - 実行中の場合は自動的にプロセスを終了
  - アンインストール前にもプロセスを自動終了
  - 完全なクリーンアップ（ユーザーデータディレクトリも削除）
  - 再インストール時のトラブル防止

## [3.2.1] - 2025-12-17

### 🔒 セキュリティ強化

- **機密情報マスキングの大幅拡充** (logging_config.py)
  - パターン数: 3種類 → 11種類に拡充
  - 新規対応: クレジットカード番号、社会保障番号、電話番号、Windowsパス内のユーザー名
  - args属性のマスキング対応追加
  - エラー時のフォールバック処理追加
  - 参考: [Better Stack - Sensitive Data](https://betterstack.com/community/guides/logging/sensitive-data/)

- **ファイル名サニタイズの強化** (pdf_converter.py)
  - 連続する無効文字の適切な処理（`folder///name` → `folder_name`）
  - 空文字列のフォールバック処理追加
  - 参考: [pathvalidate](https://github.com/thombashi/pathvalidate)

### ✨ ユーザビリティ改善

- **入力検証とフィードバックの改善** (gui/tabs/settings_tab.py)
  - 一太郎設定の入力値検証を追加（範囲チェック）
  - リトライ回数: 0～10の範囲制限
  - 保存待機時間: 5～120秒の範囲制限
  - ↓キー押下回数: 0～20の範囲制限
  - 不正な入力時に明確なエラーメッセージを表示
  - 複数のエラーを一括表示
  - 参考: [DataCamp - Input Validation](https://www.datacamp.com/tutorial/python-user-input)

### 🛠️ コード品質向上

- **インポートの最適化** (PEP 8準拠)
  - 関数内の重複インポートを削除 (pdf_converter.py:507)
  - loggingモジュールをモジュールレベルに移動 (gui/tabs/pdf_tab.py, gui/app.py)
  - 参考: [PEP 8](https://peps.python.org/pep-0008/)

- **例外処理の具体化** (gui/tabs/excel_tab.py)
  - 広範な`Exception`キャッチから具体的な例外へ変更
  - ImportError、AttributeErrorを個別にキャッチ
  - エラー原因の明確化
  - 参考: [Exception Handling Best Practices](https://medium.com/@saadjamilakhtar/5-best-practices-for-python-exception-handling-5e54b876a20)

- **マジックナンバーの定数化** (constants.py, pdf_converter.py)
  - 新規追加: `PDFConversionConstants`クラス
  - 一太郎変換の待機時間を名前付き定数に置換
  - キャンセルチェック間隔を定数化
  - 定数数: 12個の新しい定数を追加
  - 参考: [Real Python - Constants](https://realpython.com/python-constants/)

### 📚 技術的な改善

- すべての修正は2025年のPythonベストプラクティスに準拠
- 構文チェック: 全ファイル成功
- コードの保守性・可読性の向上
- デバッグ性の向上（具体的な例外メッセージ）

### 🔗 参考資料

- [PEP 8 – Style Guide for Python Code](https://peps.python.org/pep-0008/)
- [Better Stack - Logging Sensitive Data](https://betterstack.com/community/guides/logging/sensitive-data/)
- [pathvalidate - GitHub](https://github.com/thombashi/pathvalidate)
- [Real Python - Python Constants](https://realpython.com/python-constants/)
- [Medium - Exception Handling Best Practices](https://medium.com/@saadjamilakhtar/5-best-practices-for-python-exception-handling-5e54b876a20)
- [DataCamp - Python Input Validation](https://www.datacamp.com/tutorial/python-user-input)

---

## [3.2] - 2025-12-10

### 修正

- **PowerPoint変換エラーの解消**
  - `Application.Visible`プロパティ設定エラーを修正
  - `WithWindow=False`パラメータのみで対応し、環境依存の問題を解決
  - Word/Excel変換でも同様のエラーハンドリングを追加（堅牢性向上）
- **一太郎接続エラーの修正**
  - ウィンドウタイトル検索パターンを修正（`".*一太郎.*"` → ファイル名ベース）
  - 一太郎のウィンドウタイトルが `[ファイル名].jtd` 形式である問題に対応
  - 待機時間を5秒に延長（実際の起動時間に合わせて調整）
  - 正規表現の特殊文字を `re.escape()` で適切にエスケープ
  - デバッグログを追加してウィンドウ検索パターンを表示
  - より確実で環境非依存な接続を実現
- **一太郎ファイル名入力の改善**
  - 「印刷結果を名前を付けて保存」ダイアログを直接制御
  - **4段階のフォールバック機能**で確実に入力
    1. ComboBox内のEditコントロール（推奨方法）
    2. auto_id="1148"のEditコントロール
    3. pywinauto.keyboard.send_keysによるキーボード入力
    4. 従来のsend_keysによる代替入力
  - ダイアログのタイトルで確実に検出（`".*名前.*保存.*"`パターン）
  - 詳細なデバッグログで各方法の成功/失敗を記録
  - `print_control_identifiers()`でダイアログ構造を診断
  - 印刷実行後の待機時間を3秒→1秒に短縮（高速化）
  - ファイル名入力の確実性が大幅に向上

### 追加

- **一太郎変換のカスタマイズ強化**
  - `down_arrow_count`設定を追加（デフォルト: 5回）
  - **設定タブのUIで下矢印回数を直接変更可能**
  - プリンタの並び順が変わった場合も簡単に調整できる
  - わかりやすい説明とヒントを表示
  - config.jsonで下矢印キー押下回数をカスタマイズ可能に
  - プリンタ選択の柔軟性を向上
- **ログファイルへのアクセス改善**
  - 設定画面に「📄 ログファイルを開く」ボタンを追加
  - ワンクリックでログファイルを開けるように改善
  - デバッグ情報の確認が容易に

### 改善

- **一太郎変換の単純化とリトライ改善**
  - GUI自動化コードを大幅に単純化
  - 一時ファイル検出方式の問題を解決し、シンプルな時間ベース待機に変更
  - 環境非依存で確実に動作する実装に改善
  - 不要な`_wait_for_ichitaro_ready`メソッドを削除
  - 3秒待機 + pywinauto接続のシンプルなフロー
  - コードの可読性と保守性が向上
- **設定の整理**
  - 一太郎設定を4項目に集約：`ichitaro_ready_timeout`、`max_retries`、`down_arrow_count`、`save_wait_seconds`
  - 未使用の設定項目（`open_wait_seconds`、`dialog_wait_seconds`など）を削除
- **ログ出力の強化**
  - PowerPoint起動時に`WithWindow=False`使用を明示
  - Office変換時の詳細なデバッグ情報を追加
  - 一太郎接続時の2段階待機状態を明確に表示
- **テストの拡充**
  - `down_arrow_count`設定のテストケースを追加
  - デフォルト設定とカスタム設定の両方を検証
  - 新しい設定項目に対応したテストに更新

### 技術的な改善

- COMオブジェクトのエラーハンドリングを強化
- Office変換の信頼性を向上
- コードの保守性を大幅に改善
- 一太郎変換の信頼性向上（環境依存問題の解消）
- 高速環境での待機時間短縮、低速環境での確実な起動検出

---

## [3.1] - 以前のバージョン

### 機能

- PDF統合機能の実装
- Office文書（Word、Excel、PowerPoint）からPDF変換
- 一太郎文書からPDF変換
- 画像ファイルからPDF変換
- 区切りページの自動生成
- キャンセル機能の実装
- 設定の永続化

### UI

- タブベースのGUI実装
- リアルタイムログ表示
- プログレスバー表示
- ドラッグ&ドロップ対応

### システム

- Ghostscript自動検出
- 設定ファイル管理
- エラーハンドリング
- 一時ファイル管理
