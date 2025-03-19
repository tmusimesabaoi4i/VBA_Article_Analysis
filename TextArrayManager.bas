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


' サンプル
' 
        ' Sub TestFindLinesMatch()
        '     Dim manager1 As New TextArrayManager
        '     Dim manager2 As New TextArrayManager
        '     Dim sampleText1 As String
        '     Dim sampleText2 As String
        '     Dim result As Variant
            
        '     ' sampleText1 と sampleText2 の内容を設定
        '     sampleText1 = "これは1行目です。" & vbCrLf & _
        '                  "開始：完全一致行です。" & vbCrLf & _
        '                  "これは3行目です。" & vbCrLf & _
        '                  "部分一致テスト行です。" & vbCrLf & _
        '                  "開始：完全一致行です。" & vbCrLf & _
        '                  "開始：完全一致行です。" & vbCrLf & _
        '                  "終了：完全一致行です。" & vbCrLf & _
        '                  "完全一致行です。" & vbCrLf & _
            
        '     sampleText2 = "開始：完全一致行です。" & vbCrLf & _
        '                  "終了：完全一致行です。"
            
        '     ' manager1 と manager2 にテキストを設定
        '     manager1.SetOriginalText sampleText1
        '     manager2.SetOriginalText sampleText2
            
        '     ' manager1 における manager2 の位置を検索
        '     result = manager1.FindLinesMatch(manager2)
            
        '     ' 結果を表示
        '     If result(1) <> -1 Then
        '         Debug.Print "一致範囲: 開始行 = " & result(1) & " 終了行 = " & result(2)
        '     Else
        '         Debug.Print "一致する範囲は見つかりませんでした。"
        '     End If
        ' End Sub
' FindLinesMatch メソッド
' 
' このメソッドは、現在の TextArrayManager インスタンス（Me）の LinesArray と
' 指定された target の LinesArray を比較し、target の行が Me の中に完全に一致する範囲を特定します。
' 一致が見つかると、Me の中で target がどこにあるかを示す開始行と終了行を返します。
' 
' 入力: 
'   target - 比較対象となる TextArrayManager インスタンス
' 
' 出力: 
'   配列（1行目 - 開始行、2行目 - 終了行）
'   一致が見つからない場合は、-1,-1 を返します。
Public Function FindLinesMatch(target As TextArrayManager) As Variant
    Dim startLine As Long
    Dim endLine As Long
    Dim i As Long
    Dim j As Long
    Dim matchFound As Boolean
    Dim result(1 To 2) As Long ' 配列[0] - 開始行, [1] - 終了行
    
    ' 一致が見つかったかどうかを追跡するフラグ
    matchFound = False
    startLine = -1
    endLine = -1
    
    ' MeのLinesArrayとtargetのLinesArrayを比較
    ' Meの各行について、target.LinesArray(0)と一致する行を探す
    For i = LBound(LinesArray) To UBound(LinesArray)
        ' Me(i) と target の1行目を比較
        If Trim(LinesArray(i)) = Trim(target.GetLinesArray()(0)) Then
            startLine = i ' 一致した最初の行をstartLineとして設定
            ' startLine から target.LinesArray の各行を順番に比較
            For j = LBound(target.GetLinesArray()) To UBound(target.GetLinesArray())
                ' targetの行をMeの対応する行と比較
                ' Me(i + j) が target(j) と一致するか確認
                If i + j <= UBound(LinesArray) And Trim(LinesArray(i + j)) = Trim(target.GetLinesArray()(j)) Then
                    ' 一致した場合、endLineを更新
                    If j = UBound(target.GetLinesArray()) Then
                        endLine = i + j ' 最後の一致行
                        matchFound = True ' 一致が見つかったことを記録
                    End If
                Else
                    ' 一致しなかった場合、比較を終了
                    matchFound = False
                    Exit For ' targetの残りの行との一致を確認しない
                End If
            Next j
        End If
        
        ' 一致した場合、ループを抜ける
        If matchFound Then Exit For
    Next i
    
    ' 結果を格納
    ' 一致が見つかった場合、startLine と endLine を格納
    If matchFound Then
        result(1) = startLine
        result(2) = endLine
    Else
        ' 一致しなかった場合、-1をセット（範囲なし）
        result(1) = -1
        result(2) = -1
    End If
    
    ' 結果を返す
    FindLinesMatch = result
End Function


' 新しいメソッド：RemoveMatchingLines
Public Function RemoveMatchingLines(targetManager As TextArrayManager) As TextArrayManager
    ' 新しい TextArrayManager インスタンスを作成
    Dim newManager As New TextArrayManager
    ' 変数定義
    Dim i As Long
    Dim resultLines As String
    resultLines = ""

    ' currentLines と targetLines を直接取得して使用
    For i = LBound(Me.GetLinesArray()) To UBound(Me.GetLinesArray())
        ' i 番目から targetManager の行数分切り出して完全一致を確認
        If i + UBound(targetManager.GetLinesArray()) <= UBound(Me.GetLinesArray()) Then
            ' currentLines の i 番目から targetLines の長さだけ切り出し
            If Me.GetLinesSubArray(i, i + UBound(targetManager.GetLinesArray())).IsEqual(targetManager) Then
                ' 一致する場合はその部分を削除
                i = i + UBound(targetManager.GetLinesArray()) ' targetLines と一致した部分をスキップ
            Else
                ' 一致しない場合はその行を結果に追加
                If resultLines <> "" Then resultLines = resultLines & vbCrLf ' 改行を追加
                resultLines = resultLines & Me.GetLinesArray()(i)
            End If
        Else
            ' もし切り出しの範囲が無効（最後の部分が切り取れない場合）はそのまま追加
            If resultLines <> "" Then resultLines = resultLines & vbCrLf ' 改行を追加
            resultLines = resultLines & Me.GetLinesArray()(i)
        End If
    Next i
    
    ' 結果を新しい TextArrayManager にセット
    newManager.SetOriginalText resultLines
    
    ' 新しい TextArrayManager を返す
    Set RemoveMatchingLines = newManager
End Function

' 
' TextArrayManager の GetLinesSubArray メソッドを修正
Public Function GetLinesSubArray(ByVal startIdx As Long, ByVal endIdx As Long) As TextArrayManager
    ' 新しい TextArrayManager インスタンスを作成
    Dim subManager As New TextArrayManager
    Dim i As Long
    Dim subLines As String
    
    ' 範囲が有効か確認
    If startIdx < 0 Or endIdx > UBound(Me.GetLinesArray()) Then
        ' 無効な範囲の場合、Nothing を返して処理を中止
        Set GetLinesSubArray = Nothing
        Exit Function
    End If
    
    ' 範囲が有効な場合、その範囲の行を取得
    For i = startIdx To endIdx
        subLines = subLines & Me.GetLinesArray()(i) & vbCrLf ' 行を追加
    Next i
    
    ' 取得したサブ行を新しい TextArrayManager にセット
    subManager.SetOriginalText subLines
    
    ' 新しい TextArrayManager を返す
    Set GetLinesSubArray = subManager
End Function
