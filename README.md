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

### 2. GetOriginalText() As String
- **目的**:  
元のテキスト（OriginalArray）を取得します。
戻り値: String
- **使用例**:
  ```vba
Dim text As String
text = manager.GetOriginalText()
Debug.Print text
    ```

### 3. GetLinesArray() As Variant
- **目的**:  
改行で分割された文字列の配列（LinesArray）を取得します。
戻り値: Variant (文字列配列)
- **使用例**:
  ```vba
Dim lines As Variant
lines = manager.GetLinesArray()
Debug.Print lines(0)  ' 最初の行を表示
    ```

### 4. GetLineCount() As Long
- **目的**:  
テキストの行数を取得します。
戻り値: Long
- **使用例**:
  ```vba
Dim count As Long
count = manager.GetLineCount()
Debug.Print "行数: " & count
    ```

### 5. RemoveNewlines() As TextArrayManager
- **目的**:  
テキストから改行コードを削除し、改行なしの新しい TextArrayManager インスタンスを返します。
戻り値: TextArrayManager
- **使用例**:
  ```vba
Dim newManager As TextArrayManager
Set newManager = manager.RemoveNewlines()
Debug.Print newManager.GetOriginalText()
    ```

### 6. CopyFrom(ByVal sourceManager As TextArrayManager)
- **目的**:  
指定した別の TextArrayManager インスタンスからテキストと行情報をコピーします。
- **パラメータ**:
sourceManager (TextArrayManager): コピー元のインスタンス
- **使用例**:
  ```vba
Dim manager1 As New TextArrayManager
Dim manager2 As New TextArrayManager
manager1.SetOriginalText "サンプルテキスト"
manager2.CopyFrom manager1
    ```

### 7. RemoveSpacesAndTabs()
- **目的**:  
テキスト内の半角・全角スペース、タブ、垂直タブを削除し、改めて LinesArray と LineCount を更新します。
- **使用例**:
  ```vba
manager.RemoveSpacesAndTabs()
Debug.Print manager.GetOriginalText()
    ```

### 8. ConvertToHalfWidth()
- **目的**:  
文字列内の全角文字を半角に変換し、LinesArray と LineCount を更新します。
※ 内部では StrConv 関数を使用します。
- **使用例**:
  ```vba
manager.ConvertToHalfWidth()
Debug.Print manager.GetOriginalText()
    ```
### 9. ExtractLines(ByVal target As String) As TextArrayManager
- **目的**:  
各行の先頭が指定した文字列 (target) と一致する行だけを抽出し、新しい TextArrayManager インスタンスとして返します。
- **パラメータ**:
target (String): 抽出対象の行の先頭文字列
戻り値: TextArrayManager
- **使用例**:
  ```vba
Dim extractedManager As TextArrayManager
Set extractedManager = manager.ExtractLines("特定の行")
Debug.Print extractedManager.GetOriginalText()
    ```
### 10. ExtractTextBetweenTargets(ByVal targetStart As String, ByVal targetEnd As String) As TextArrayManager
- **目的**:  
指定された開始文字列 (targetStart) と終了文字列 (targetEnd) の間にある行を抽出し、新しい TextArrayManager インスタンスとして返します。
※ 開始ターゲット行自体は抽出されず、終了ターゲットに到達した時点で抽出を終了します。
- **パラメータ**:
targetStart (String): 抽出開始の目印となる文字列
targetEnd (String): 抽出終了の目印となる文字列
戻り値: TextArrayManager
- **使用例**:
  ```vba
Dim betweenManager As TextArrayManager
Set betweenManager = manager.ExtractTextBetweenTargets("開始ターゲット", "終了ターゲット")
Debug.Print betweenManager.GetOriginalText()
    ```
### 11. RemoveLineByExactMatch(ByVal target As String) As TextArrayManager
- **目的**:  
テキスト内で指定した文字列と完全一致する行を削除し、新しい TextArrayManager インスタンスとして返します。
※ ※ ※ 注意: コード中に LinesArray の参照でタイポ（inesArray になっている）があるため、実装前に修正が必要です。
- **パラメータ**:
target (String): 削除対象の行と完全一致する文字列
戻り値: TextArrayManager
- **使用例**:
  ```vba
Dim filteredManager As TextArrayManager
Set filteredManager = manager.RemoveLineByExactMatch("削除したい行です。")
Debug.Print filteredManager.GetOriginalText()
    ```
### 12. IsEqual(ByVal otherManager As TextArrayManager) As Boolean
- **目的**:  
他の TextArrayManager インスタンスと、元のテキスト (OriginalArray) および各行 (LinesArray) が一致するかを判定します。
- **パラメータ**:
otherManager (TextArrayManager): 比較対象のインスタンス
戻り値: Boolean
一致すれば True、そうでなければ False
- **使用例**:
  ```vba
Dim result As Boolean
result = manager.IsEqual(anotherManager)
Debug.Print "一致: " & result
    ```
### 13. RemoveMatchingText(ByVal target As TextArrayManager) As TextArrayManager
- **目的**:  
引数で指定された TextArrayManager インスタンスのテキストと一致する部分を、元のテキストから削除し、削除後の新しい TextArrayManager インスタンスを返します。
※ このメソッドは最初に一致した部分のみ置換します。
- **パラメータ**:
target (TextArrayManager): 削除対象のテキストを保持するインスタンス
戻り値: TextArrayManager
- **使用例**:
  ```vba
Dim reducedManager As TextArrayManager
Set reducedManager = manager.RemoveMatchingText(targetManager)
Debug.Print reducedManager.GetOriginalText()
使用例（サンプルコード）
以下は、各メソッドの利用例として参考にしてください。

vba

Sub TestTextArrayManager()
    Dim manager As New TextArrayManager
    Dim sampleText As String
    Dim i As Long
    
    sampleText = "これは1行目です。" & vbCrLf & _
                 "開始ターゲット" & vbCrLf & _
                 "抽出する内容1" & vbCrLf & _
                 "抽出する内容2" & vbCrLf & _
                 "終了ターゲット" & vbCrLf & _
                 "これは別の行です。"
    
    ' テキストを設定
    manager.SetOriginalText sampleText
    
    ' 特定の行を抽出する例
    Dim extractedManager As TextArrayManager
    Set extractedManager = manager.ExtractLines("開始ターゲット")
    
    ' 結果を表示
    For i = 0 To extractedManager.GetLineCount() - 1
        Debug.Print extractedManager.GetLinesArray()(i)
    Next i
End Sub
    ```
