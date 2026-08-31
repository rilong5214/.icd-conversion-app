# ICD Converter

iCAD SXの付属変換プログラムを、Windows 11に合わせたデスクトップUIから実行するファイル変換・ランチャーアプリです。

UIと変換制御はC#、Avalonia 12.1、.NET 8で実装されています。配布されている`ICDConverter.exe`はWindows x64向けの自己完結型単一EXEで、Python、`EXE.bat`、別途インストールした.NETランタイムは必要ありません。

## 動作条件

- 対応OS: Windows 10 / Windows 11（64ビット）
- iCAD SXがインストールされていること
- PDF変換を使用する場合は、PDFCreatorのAutoSave設定が完了していること

`ICDConverter.exe`にはiCAD SX本体や、iCAD SX付属の変換プログラムは含まれていません。変換を実行するPCに、正規のiCAD SX環境が必要です。

## 起動方法

1. このリポジトリから`ICDConverter.exe`をダウンロードします。
2. `ICDConverter.exe`をダブルクリックします。
3. 初回起動後、左側の`設定`を開いてiCAD SXの検出状態を確認します。
4. `ファイル変換`画面へファイルを追加し、`変換開始`を押します。

署名されていないEXEのため、Windows Defender SmartScreenに発行元不明の警告が表示される場合があります。

## 対応する変換

| 変換モード | 入力 | 出力 |
|---|---|---|
| ICD → STEP | `.icd` | STEP |
| ICD → Parasolid | `.icd` | `.x_b` / `.x_t` / `.xmt_bin` / `.xmt_txt` |
| STEP / Parasolid → ICD | `.stp` / `.step` / `.x_b` / `.x_t` / `.xmt_bin` / `.xmt_txt` | `.icd` |
| CATIA → ICD | `.catpart` / `.catproduct` | `.icd` |
| ICD → DWG / DXF / DXB | `.icd` | `.dwg` / `.dxf` / `.dxb` |
| DWG / DXF → ICD | `.dwg` / `.dxf` | `.icd` |
| ICD → PDF | `.icd` | `.pdf` |

ファイルをドラッグ・アンド・ドロップすると、拡張子から利用可能な変換モードを自動判定します。同じ拡張子に複数の変換候補がある場合は、その候補だけが選択肢として表示されます。

## iCAD SXの検出

アプリは次の順番でiCAD SXのインストールフォルダーを検索します。

1. 設定画面で保存したフォルダー
2. 環境変数`ICADDIR`
3. `C:\ICADSX`

必要な変換プログラムが見つからない場合、該当モードは変換候補に表示されません。`設定`画面には、不足しているEXE名と利用可能なモード数が表示されます。

## 変換オプション

- STEP精度: `0（標準精度）` / `2（高精度・推奨）`
- Parasolidバージョン: 最新自動選択、R20からR29
- Parasolid出力形式: AUTO、`.x_b`、`.x_t`、`.xmt_bin`、`.xmt_txt`
- 2D出力形式とCADバージョン
- PDFプロッター番号とPDFCreator AutoSaveフォルダー
- 出力先フォルダー

詳細オプションとiCAD SXフォルダーは`%LOCALAPPDATA%\ICDConverter\settings.json`へ保存されます。ウィンドウを閉じた後やWindows再起動後も設定は保持されます。

## ランチャー

よく使用するファイルやフォルダーを四角いタイルとして登録できます。タイルをクリックすると、Windowsの既定アプリまたはエクスプローラーで対象を開きます。

## PDF変換

`ICD → PDF`は、iCAD SXの`SXPLOT.exe`とPDFCreatorを使用します。

1. PDFCreatorでAutoSaveを有効にします。
2. PDFCreatorの保存先と、アプリの`AutoSave先`を同じフォルダーに設定します。
3. PDFCreatorを割り当てたiCAD SXのプロッター番号を指定します。

アプリはSXPLOT実行後に新しいPDFを最大20秒待機し、指定した出力先へ移動します。

## トラブルシューティング

- 変換モードが表示されない: `設定`画面でiCAD SXフォルダーと不足EXEを確認してください。
- ドラッグ・アンド・ドロップできない: エクスプローラーとアプリの実行権限を揃えるか、`ファイルを選択`を使用してください。
- PDFが作成されない: PDFCreatorのAutoSave設定、保存先、プロッター番号を確認してください。
- 変換に失敗する: `詳細ログ`に表示される標準エラー出力とiCAD SXの終了コードを確認してください。

## 配布ファイル

- ファイル: `ICDConverter.exe`
- サイズ: 45.03 MiB
- 対象: Windows x64
- 形式: .NET 8 self-contained / single-file
- SHA-256: `63BF07F6FABE6DDE1368A28C292E858F65B098AEC20FC88599E5D6081AD91EC4`

不具合や改善要望は[Issues](https://github.com/rilong5214/.icd-conversion-app/issues)へ登録してください。
