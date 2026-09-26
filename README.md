# AutomaticTestPrinting

特訓レポートから確認テストの作成依頼を読み取り、既存のExcelマクロとPDF教材を使って問題・解答を準備するWindowsデスクトップアプリです。

## 現在の範囲

初期版では以下を実装しています。

- OneDrive内の教材フォルダー設定
- Google Drive（マイドライブ）の未処理フォルダーからレポートPDFを一括読込
- 同期直後・空ファイル・重複ファイルの除外
- 作成結果の出力フォルダー設定
- 複数の特訓レポートPDF選択
- 選択内容の事前検査
- 設定のローカル保存
- Windows標準OCRによるレポート読み取り（追加料金なし）
- 0度・180度の自動向き判定
- 生徒名、通常テスト依頼、段階突破テスト依頼の確認表示
- OCR座標を使った通常テストの教材名・範囲・問題数の列解析
- 読み取り結果から原本PDFを開く確認導線
- OCRで読み取った生徒名・教材名・開始番号・終了番号・問題数の画面修正と即時再検査
- レポート単位・テスト単位のスキップ、エラー・未対応教材の一括スキップ
- 要確認項目だけの絞り込み表示
- 大量レポート向けの2件並行OCRと進捗件数表示
- ターゲット1900（6訂版）、英検準1級単熟語EX（第2版）、英熟語ターゲット1000（5訂版）、Vintage（4th Edition）、国語力を伸ばす 語彙1700（シグマベスト）、速読英熟語（新版・改訂版）、共通テスト 公共・政治経済 集中講義（五訂版）のExcel連携
- 教材名の空白・括弧・全角半角などの表記ゆれを吸収
- 範囲・問題数・ファイル存在の事前検査（Vintageは最大25問、速読英熟語はテーマ番号指定・最大50問）
- 問題番号を基準に原本の問題解答リストを照合し、問題文と解答の取り違えを防止
- 元のExcelを変更せず、対応教材の問題PDF・解答PDFを個別に作成
- 原本確認チェックと作成前の最終確認（実プリンターへの送信なし）
- 既存教材と同じExcel形式の教材を画面から追加登録

## 教材設定

対応教材は `materials.json` で管理します。
同じExcel形式の教材は、画面右上の「教材を登録」から追加できます。
利用者がJSONを直接編集する必要はありません。登録時にはマクロを無効にした状態で
Excelを読み取り専用で開き、選択した既存教材と同じシート構成かを確認します。

追加設定は `%LocalAppData%\AutomaticTestPrinting\materials.json` に保存され、
2回目以降の保存時には直前の設定を `materials.json.backup` に退避します。
配布版を更新しても、追加登録した教材は引き継がれます。

設定ファイルには、教材名の候補、Excelファイル名の一部、最大範囲・問題数、
シート名、範囲入力セル、PDFレイアウトなどを記録します。起動後の読み取り時に
必須項目、数値、教材IDの重複を検査し、不正な設定は日本語のエラーとして表示します。

既存教材とシート構成や動作が異なるExcelは、新しい形式の解析と処理エンジン追加が必要です。

PDF教材の抽出と実プリンターへの印刷は今後の段階で追加します。

OCR結果は誤認識の可能性があるため、印刷処理を接続するまでは原本PDFとの確認を必須とします。

## 技術構成

- C# / .NET 10 LTS
- WPF
- Windows 11 x64
- Excel 2021 64bit
- xUnit

## プロジェクト構成

```text
src/AutomaticTestPrinting.App           WPFアプリ
src/AutomaticTestPrinting.Core          共通モデル・検査・設定保存
tests/AutomaticTestPrinting.Core.Tests  自動テスト
```

## 開発

```powershell
dotnet restore
dotnet build
dotnet test
dotnet run --project src/AutomaticTestPrinting.App
```

## Windows試験配布版の作成

Windows 11 64bit向けの自己完結型アプリをZIPにまとめます。.NETの別途インストールは不要です。

```powershell
powershell -ExecutionPolicy Bypass -File scripts/Publish-Distribution.ps1
```

成果物は既定で `artifacts/distribution` に作成されます。配布版には教材、レポート、生徒情報を含めません。

教材、レポート、生徒情報を含むファイルはGitHubへ追加しないでください。
