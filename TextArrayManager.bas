' クラス名: TextArrayManager
' このクラスは、文字列配列を管理し、特定の行を抽出するメソッドも提供します。

Private OriginalArray As String
Private LinesArray() As String
Private LineCount As Long

' 文字列配列を設定するメソッド
Public Sub SetOriginalText(ByVal text As String)
    Me.OriginalArray = text
    ' 改行で分割してLinesArrayに格納
    Me.LinesArray = Split(Me.OriginalArray, vbCrLf)
    ' 行数を設定
    Me.LineCount = UBound(LinesArray) - LBound(LinesArray) + 1
End Sub

' 文字列配列を取得するメソッド
Public Function GetOriginalText() As String
    Set GetOriginalText = Me.OriginalArray
End Function

' 改行で分割された配列を取得するメソッド
Public Function GetLinesArray() As Variant
    Set GetLinesArray = Me.LinesArray
End Function

' 行数を取得するメソッド
Public Function GetLineCount() As Long
    Set GetLineCount = Me.LineCount
End Function


' サンプル
' 
        ' Sub TestTextArrayManager()
        '     Dim originalManager As New TextArrayManager
        '     Dim modifiedManager As TextArrayManager
        '     Dim sampleText As String
        '     Dim i As Long
            
        '     sampleText = "これは1行目です。" & vbCrLf & "これは2行目。" & vbCrLf & "3行目もあります。"
            
        '     ' オリジナルのテキストをセット
        '     originalManager.SetOriginalText sampleText
            
        '     ' 改行を削除した新しいインスタンスを取得
        '     Set modifiedManager = originalManager.RemoveNewlines
            
        '     ' 新しいインスタンスで変更後のテキストを表示
        '     For i = 0 To modifiedManager.GetLineCount - 1
        '         Debug.Print modifiedManager.GetLinesArray()(i)
        '     Next i
        ' End Sub

' 改行を削除するメソッド（TextArrayManager型を返す）
Public Function RemoveNewlines() As TextArrayManager
    Dim newManager As New TextArrayManager
    ' 改行を削除
    newManager.SetOriginalText Replace(Replace(Replace(Replace(Me.OriginalArray, vbNewLine, ""),vbLf,""),vbCr,""),vbCrLf,"")
    ' 削除後の新しいTextArrayManagerインスタンスを返す
    Set RemoveNewlines = newManager
End Function


' サンプル 
' 
        ' Sub TestTextArrayManager()
        '     Dim manager1 As New TextArrayManager
        '     Dim manager2 As New TextArrayManager
        '     Dim sampleText As String
            
        '     sampleText = "これは1行目です。" & vbCrLf & _
        '                  "これは2行目です。" & vbCrLf & _
        '                  "これは3行目です。"
            
        '     ' manager1 にテキストを設定
        '     manager1.SetOriginalText sampleText
            
        '     ' manager2 に manager1 の内容をコピー
        '     manager2.CopyFrom manager1
            
        '     ' manager2 の内容を表示
        '     Dim i As Long
        '     For i = 0 To manager2.GetLineCount - 1
        '         Debug.Print manager2.GetLinesArray()(i)
        '     Next i
        ' End Sub

' TextArrayManager 型を代入するメソッド
Public Sub CopyFrom(ByVal sourceManager As TextArrayManager)
    ' 元のTextArrayManagerインスタンスからデータをコピー
    Me.SetOriginalText sourceManager.GetOriginalText()
    ' 新しいLinesArrayと行数を再計算
    Me.LinesArray = sourceManager.GetLinesArray()
    Me.LineCount = UBound(Me.LinesArray) - LBound(Me.LinesArray) + 1
End Sub


' サンプル
' 
        ' Sub TestTextArrayManager()
        '     Dim manager As New TextArrayManager
        '     Dim sampleText As String
        '     Dim i As Long
            
        '     sampleText = "これは１行目です。" & vbCrLf & "１２３４５６７８９０" & vbCrLf & "ＡＢＣＤＥＦ"
            
        '     ' オリジナルのテキストをセット
        '     manager.SetOriginalText sampleText
            
        '     ' 全てを半角に変換
        '     manager.RemoveSpacesAndTabs
        '     manager.ConvertToHalfWidth
            
        '     ' 半角に変換した後のテキストを表示
        '     For i = 0 To manager.GetLineCount - 1
        '         Debug.Print manager.GetLinesArray()(i)
        '     Next i
        ' End Sub

' スペースとタブを削除するメソッド
Public Sub RemoveSpacesAndTabs()
    ' 文字列内のタブとスペースを削除
    Me.OriginalArray = Replace(Me.OriginalArray, " ", "") ' スペースを削除
    Me.OriginalArray = Replace(Me.OriginalArray, "　", "") ' スペースを削除
    Me.OriginalArray = Replace(Me.OriginalArray, vbTab, "") ' タブを削除
    Me.OriginalArray = Replace(Me.OriginalArray, vbVerticalTab, "") ' タブを削除

    ' スペースとタブを削除後、改めて配列を作り直す
    Me.LinesArray = Split(Me.OriginalArray, vbCrLf)
    Me.LineCount = UBound(Me.LinesArray) - LBound(Me.LinesArray) + 1
End Sub

' 全て半角に変換するメソッド
Public Sub ConvertToHalfWidth()
    ' 文字列内の全角文字を半角に変換
    Me.OriginalArray = StrConv(Me.OriginalArray, vbNarrow)
    
    ' 半角に変換後、改めて配列を作り直す
    Me.LinesArray = Split(Me.OriginalArray, vbCrLf)
    Me.LineCount = UBound(Me.LinesArray) - LBound(Me.LinesArray) + 1
End Sub


' サンプル
' 
        ' Sub TestTextArrayManager()
        '     Dim manager As New TextArrayManager
        '     Dim extractedManager As TextArrayManager
        '     Dim sampleText As String
        '     Dim i As Long
            
        '     sampleText = "これは1行目です。" & vbCrLf & "特定の行です。" & vbCrLf & "これは別の行です。" & vbCrLf & "特定の行が続きます。"
            
        '     ' テキストをセット
        '     manager.SetOriginalText sampleText
            
        '     ' 「特定の行」で始まる行を抽出
        '     Set extractedManager = manager.ExtractLines("特定の行")
            
        '     ' 抽出した行を表示
        '     For i = 0 To extractedManager.GetLineCount - 1
        '         Debug.Print extractedManager.GetLinesArray()(i)
        '     Next i
        ' End Sub

' 行を抽出するメソッド（TextArrayManager型を返す）
Public Function ExtractLines(ByVal target As String) As TextArrayManager
    Dim result As String
    Dim i As Long
    Dim newManager As New TextArrayManager
    
    ' 結果を初期化
    result = ""
    
    ' 各行を確認してtargetで始まる行を抽出
    For i = LBound(Me.LinesArray) To UBound(Me.LinesArray)
        If Trim(Me.LinesArray(i)) Like target & "*" Then
            result = result & Me.LinesArray(i) & vbCrLf
        End If
    Next i
    
    ' 最後の改行を除去
    If Len(result) > 0 Then
        result = Left(result, Len(result) - Len(vbCrLf))
    End If
    
    ' 結果を新しいTextArrayManagerインスタンスにセット
    newManager.SetOriginalText result
    
    ' 新しいTextArrayManagerインスタンスを返す
    Set ExtractLines = newManager
End Function


' サンプル
' 
        ' Sub TestTextArrayManager()
        '     Dim manager As New TextArrayManager
        '     Dim extractedManager As TextArrayManager
        '     Dim sampleText As String
        '     Dim i As Long
            
        '     sampleText = "これは1行目です。" & vbCrLf & _
        '                  "開始ターゲット" & vbCrLf & _
        '                  "抽出する内容1" & vbCrLf & _
        '                  "抽出する内容2" & vbCrLf & _
        '                  "終了ターゲット" & vbCrLf & _
        '                  "これは別の行です。"
            
        '     ' テキストをセット
        '     manager.SetOriginalText sampleText
            
        '     ' 「開始ターゲット」から「終了ターゲット」までの行を抽出
        '     Set extractedManager = manager.ExtractTextBetweenTargets("開始ターゲット", "終了ターゲット")
            
        '     ' 抽出した行を表示
        '     For i = 0 To extractedManager.GetLineCount - 1
        '         Debug.Print extractedManager.GetLinesArray()(i)
        '     Next i
        ' End Sub

' 二つの文章間の文章を抽出するメソッド（TextArrayManager型を返す）
Public Function ExtractTextBetweenTargets(ByVal targetStart As String, ByVal targetEnd As String) As TextArrayManager
    Dim result As String
    Dim i As Long
    Dim insideTarget As Boolean
    Dim newManager As New TextArrayManager
    
    ' 結果を初期化
    result = ""
    insideTarget = False
    
    ' 各行を確認してtargetStartとtargetEndの間の行を抽出
    For i = LBound(Me.LinesArray) To UBound(Me.LinesArray)
        If Trim(Me.LinesArray(i)) = targetStart Then
            insideTarget = True
        End If
        
        If Trim(Me.LinesArray(i)) = targetEnd Then
            insideTarget = False
            Exit For
        End If
        
        If insideTarget Then
            result = result & LinesArray(i) & vbCrLf
        End If
    Next i
    
    ' 最後の改行を除去
    If Len(result) > 0 Then
        result = Left(result, Len(result) - Len(vbCrLf))
    End If
    
    ' 結果を新しいTextArrayManagerインスタンスにセット
    newManager.SetOriginalText result
    
    ' 新しいTextArrayManagerインスタンスを返す
    Set ExtractTextBetweenTargets = newManager
End Function


' サンプル
' 
        ' Sub TestRemoveLineByExactMatch()
        '     Dim manager As New TextArrayManager
        '     Dim newManager As TextArrayManager
        '     Dim sampleText As String
        '     Dim i As Long
            
        '     sampleText = "これは1行目です。" & vbCrLf & _
        '                 "削除したい行です。" & vbCrLf & _
        '                 "これは3行目です。" & vbCrLf & _
        '                 "削除したい行です。" & vbCrLf
            
        '     ' manager にテキストを設定
        '     manager.SetOriginalText sampleText
            
        '     ' "削除したい行です。" の行を削除
        '     Set newManager = manager.RemoveLineByExactMatch("削除したい行です。")
            
        '     ' 結果を表示
        '     For i = 0 To newManager.GetLineCount - 1
        '         Debug.Print newManager.GetLinesArray()(i)
        '     Next i
        ' End Sub

' TextArrayManager 型に指定した文字列と同じ文字列の行を削除するメソッド（TextArrayManager 型を返す）
Public Function RemoveLineByExactMatch(ByVal target As String) As TextArrayManager
    Dim newManager As New TextArrayManager
    Dim result As String
    Dim i As Long
    
    ' 結果を初期化
    result = ""
    
    ' 各行を確認してtargetと一致する行を削除
    For i = LBound(Me.inesArray) To UBound(Me.LinesArray)
        If Trim(Me.LinesArray(i)) <> target Then ' 一致する行はスキップ
            result = result & Me.LinesArray(i) & vbCrLf
        End If
    Next i
    
    ' 最後の改行を除去
    If Len(result) > 0 Then
        result = Left(result, Len(result) - Len(vbCrLf))
    End If
    
    ' 新しいTextArrayManagerインスタンスに結果をセット
    newManager.SetOriginalText result
    
    ' 新しいTextArrayManagerインスタンスを返す
    Set RemoveLineByExactMatch = newManager
End Function


' サンプル
' 
        ' Sub TestIsEqual()
        '     Dim manager1 As New TextArrayManager
        '     Dim manager2 As New TextArrayManager
        '     Dim manager3 As New TextArrayManager
        '     Dim sampleText1 As String
        '     Dim sampleText2 As String
        '     Dim result As Boolean
            
        '     sampleText1 = "これは1行目です。" & vbCrLf & _
        '                  "これは2行目です。" & vbCrLf & _
        '                  "これは3行目です。"
            
        '     sampleText2 = "これは1行目です。" & vbCrLf & _
        '                  "これは2行目です。" & vbCrLf & _
        '                  "これは3行目です。"
            
        '     ' manager1 と manager2 に同じテキストを設定
        '     manager1.SetOriginalText sampleText1
        '     manager2.SetOriginalText sampleText2
            
        '     ' manager3 に異なるテキストを設定
        '     manager3.SetOriginalText "異なるテキストです。" & vbCrLf & "他の行です。"
            
        '     ' manager1 と manager2 が一致するか判定
        '     result = manager1.IsEqual(manager2)
        '     Debug.Print "manager1 と manager2 は一致するか？ " & result ' True が出力されるべき
            
        '     ' manager1 と manager3 が一致するか判定
        '     result = manager1.IsEqual(manager3)
        '     Debug.Print "manager1 と manager3 は一致するか？ " & result ' False が出力されるべき
        ' End Sub

' TextArrayManager 型の2つのインスタンスが一致するか判定するメソッド（Boolean 型を返す）
Public Function IsEqual(ByVal otherManager As TextArrayManager) As Boolean
    ' OriginalArray を比較
    If Me.GetOriginalText() <> otherManager.GetOriginalText() Then
        Set IsEqual = False
        Exit Function
    End If
    
    ' LinesArray を比較
    Dim i As Long
    Dim lines1() As String
    Dim lines2() As String
    
    lines1 = Me.GetLinesArray()
    lines2 = otherManager.GetLinesArray()
    
    If UBound(lines1) <> UBound(lines2) Then
        Set IsEqual = False
        Exit Function
    End If
    
    ' 各行を比較
    For i = LBound(lines1) To UBound(lines1)
        If lines1(i) <> lines2(i) Then
            Set IsEqual = False
            Exit Function
        End If
    Next i
    
    ' すべて一致する場合は True を返す
    Set IsEqual = True
End Function

' TextArrayManager型で、一致したテキストを削除するメソッド
Public Sub RemoveMatchingText(ByVal target As TextArrayManager) As TextArrayManager
    Dim newManager As New TextArrayManager
    Dim result As String
    Dim targetText As String
    Dim pos As Long

    ' targetオブジェクトからテキストを取得
    targetText = target.GetOriginalText()

    ' 最初の一致位置を検索
    pos = InStr(OriginalArray, targetText)
    
    ' 最初に一致した部分だけ置換
    If pos > 0 Then
        result = Left(OriginalArray, pos - 1) & Mid(OriginalArray, pos + Len(targetText))
    End If

    ' 新しいTextArrayManagerインスタンスに結果をセット
    newManager.SetOriginalText result
    
    ' 新しいTextArrayManagerインスタンスを返す
    Set RemoveMatchingText = newManager
End Function
