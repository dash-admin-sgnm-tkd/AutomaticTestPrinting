# 処理中の待機表示

- レポート読み取り・教材の照合・個別PDF作成・一括PDF結合中に、待機アニメーションと処理状況を表示します。
- 件数が分かる工程では完了件数と進捗バーを表示します。進捗は残り時間の推定ではありません。
- 処理中はメイン画面の入力とボタン操作を無効にし、設定変更や二重実行を防ぎます。最小化は可能です。
- 閉じる操作では途中終了せず、「処理が完了したら自動で閉じる」か、そのまま待機するかを確認します。設定を保存してから終了します。
- アプリ内の終了操作を対象としています。タスクマネージャーからの強制終了やPCの電源断を防ぐ機能ではありません。
- Excel処理が戻らなくなった場合も、自動的な強制終了は行いません。

## 検証

`tools/BusyUiPreview` は実データやExcel、プリンター、保存済み設定を操作せず、WPFの待機表示、進捗、回転アニメーション、操作ロックと解除、処理中の終了抑止を確認してPNGを出力します。

```powershell
dotnet build tools/BusyUiPreview/BusyUiPreview.csproj --no-restore -o bin/busy-ui-preview-build
dotnet bin/busy-ui-preview-build/BusyUiPreview.dll ../output/busy-ui-preview.png
```

実際の大量レポート読み取りや実印刷による通しテストは別途行ってください。
