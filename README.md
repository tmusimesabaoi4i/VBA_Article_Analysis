# VBA_Article_Analysis
|モジュール名|説明|
| ---- | ---- |
|modFileReader|ファイルの読み込み処理 (テキストファイルのオープン、読み取り、エラーハンドリング)|
|modTextProcessor|テキストのパターン抽出、数値と文字の処理 (ソート、マージ)|
|modResultWriter|ワークシートへの結果出力|
|modMainController|上記3つのモジュールを順に実行し、エラー処理や制御フローを担う|
|modDebugOutput|デバッグ出力専用のモジュール|


# TextArrayManager クラス ドキュメント

## 概要
**TextArrayManager** クラスは、文字列全体および改行区切りの各行を管理するためのクラスです。  
主な機能としては、テキストの設定、行ごとの分割、特定条件に基づく行の抽出や削除、文字列のフォーマット変換（半角変換、スペースやタブの削除）などを提供します。

## 内部変数（プライベートメンバ）
- **OriginalArray**  
  - 型: String  
  - 説明: 元のテキスト全体を保持する変数
- **LinesArray**  
  - 型: 配列(String)  
  - 説明: `OriginalArray` を改行で分割した各行を格納する配列
- **LineCount**  
  - 型: Long  
  - 説明: `LinesArray` に格納された行数

## メソッド一覧

### 1. SetOriginalText(ByVal text As String)
- **目的**:  
  指定されたテキストを `OriginalArray` にセットし、改行で分割して `LinesArray` に格納、さらに行数を計算します。
- **パラメータ**:
  - `text` (String): 設定するテキスト
- **使用例**:
  ```vba
  Dim manager As New TextArrayManager
  manager.SetOriginalText "これは1行目です。" & vbCrLf & "これは2行目です。"
```
