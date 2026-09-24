# AutomaticTestPrinting

特訓レポートから確認テストの作成依頼を読み取り、既存のExcelマクロとPDF教材を使って問題・解答を準備するWindowsデスクトップアプリです。

## 現在の範囲

初期版では以下を実装しています。

- OneDrive内の教材フォルダー設定
- 作成結果の出力フォルダー設定
- 複数の特訓レポートPDF選択
- 選択内容の事前検査
- 設定のローカル保存
- Windows標準OCRによるレポート読み取り（追加料金なし）
- 0度・180度の自動向き判定
- 生徒名、通常テスト依頼、段階突破テスト依頼の確認表示
- OCR座標を使った通常テストの教材名・範囲・問題数の列解析
- 読み取り結果から原本PDFを開く確認導線

Excelマクロ連携、PDF抽出、印刷は今後の段階で追加します。

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

教材、レポート、生徒情報を含むファイルはGitHubへ追加しないでください。
