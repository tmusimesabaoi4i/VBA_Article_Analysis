' メインの処理 (パブリック)
Public Sub ProcessText()
    Dim tmp_bk As String
    Dim tmp_dot As String
    Dim tmp_bk_line() As String
    Dim tmp_dot_line() As String
    tmp_bk = ExtractLines(LoadedText, "●")
    tmp_dot = ExtractLines(LoadedText, "・")

    ReDim tmp_bk_line(count_vbCrLf(tmp_bk))
    ReDim tmp_dot_line(count_vbCrLf(tmp_dot))

    tmp_bk_line = Extract_text_BlackCircle(LoadedText)
    tmp_dot_line = Extract_text_Dots(tmp_bk_line(0))
End Sub

' パターン抽出 (プライベート)
Private Function ExtractPatterns(ByVal text As String) As Collection

End Function

' 数値抽出 (プライベート)
Private Function ExtractNumbers(ByVal text As String) As Object ' Dictionary型

End Function

' 数値のソートとマージ (プライベート)
Private Function SortAndMergeNumbers(ByRef numbers As Object) As Collection

End Function

' 行を抽出する関数
Private Function ExtractLines(ExtractLines_input_A As String, ExtractLines_target As String) As String
    Dim ExtractLines_Split_Lines() As String
    Dim Result As String
    Dim ExtractLines_i As Long
    
    ' 改行コードで文字列を分割
    ExtractLines_Split_Lines = Split(ExtractLines_input_A, vbCrLf)
    
    ' 抽出結果を初期化
    Result = ""
    
    ' 各行を確認してExtractLines_targetで始まる行を抽出
    For ExtractLines_i = LBound(ExtractLines_Split_Lines) To UBound(ExtractLines_Split_Lines)
        If Trim(ExtractLines_Split_Lines(ExtractLines_i)) Like ExtractLines_target & "*" Then
            Result = Result & ExtractLines_Split_Lines(ExtractLines_i) & vbCrLf
        End If
    Next ExtractLines_i
    
    ' 最後の改行を除去
    If Len(Result) > 0 Then
        Result = Left(Result, Len(Result) - Len(vbCrLf))
    End If
    
    ' 結果を返す
    ExtractLines = Result
End Function

' 二つの文章間の文章を抽出
Private Function ExtractTextBetweenTargets(ExtractTextBetweenTargets_input As String, ExtractTextBetweenTargets_target_start As String, ExtractTextBetweenTargets_target_end As String) As String
    Dim ExtractTextBetweenTargets_Split_Lines() As String
    Dim Result As String
    Dim insideTarget As Boolean
    Dim i As Long

    ' 改行コードで文字列を分割
    ExtractTextBetweenTargets_Split_Lines = Split(ExtractTextBetweenTargets_input, vbCrLf)

    ' 抽出結果を初期化
    Result = ""
    insideTarget = False

    ' 各行を確認してExtractTextBetweenTargets_target_startとExtractTextBetweenTargets_target_endの間の行を抽出
    For i = LBound(ExtractTextBetweenTargets_Split_Lines) To UBound(ExtractTextBetweenTargets_Split_Lines)
        If Trim(ExtractTextBetweenTargets_Split_Lines(i)) = ExtractTextBetweenTargets_target_start Then
            insideTarget = True
        End If
        
        If Trim(ExtractTextBetweenTargets_Split_Lines(i)) = ExtractTextBetweenTargets_target_end Then
            insideTarget = False
            Exit For
        End If
        
        If insideTarget Then
            Result = Result & ExtractTextBetweenTargets_Split_Lines(i) & vbCrLf
        End If
    Next i
    
    ' 最後の改行を除去
    If Len(Result) > 0 Then
        Result = Left(Result, Len(Result) - Len(vbCrLf))
    End If

    ExtractTextBetweenTargets = Result
End Function

' 行数をカウント
Private Function count_vbCrLf(count_vbCrLf_input As String) As Long
  count_vbCrLf = UBound(Split(count_vbCrLf_input, vbCrLf)) + 1
End Function


' ###########################################################################################
        ' 文章中の●から●までの文章を抽出する関数

        ' ●　XXXXXXXXXXXX
        ' YYYYYYYYYYY
        ' YYYYYYYYYYY
        ' YYYYYYYYYYY
        ' YYYYYYYYYYY

        ' ●　XXXXXXXXXXXX
        ' ZZZZZZZZZZZ
        ' ZZZZZZZZZZZ
        ' ZZZZZZZZZZZ
        ' ZZZZZZZZZZZ

        ' という入力があったときに、

        ' 一つ目が
        ' YYYYYYYYYYY
        ' YYYYYYYYYYY
        ' YYYYYYYYYYY
        ' YYYYYYYYYYY

        ' 二つ目が
        ' ZZZZZZZZZZZ
        ' ZZZZZZZZZZZ
        ' ZZZZZZZZZZZ
        ' ZZZZZZZZZZZ

        ' となる。
'
Private Function Extract_text_BlackCircle(Extract_text_BlackCircle_input As String)
    Dim BlackCircle As String
    Dim BlackCircle_Lines() As String
    Dim Result_BlackCircle_Lines() As String
    Dim Extract_text_BlackCircle_i As Long

    ' ●の列を抽出
    BlackCircle = ExtractLines(Extract_text_BlackCircle_input, "●")

    ' ●の配列を生成「絶対に存在しない●●●●●という文字列を加えている」
    BlackCircle = BlackCircle & vbCrLf & "●●●●●"

    ' 配列化している
    BlackCircle_Lines = Split(BlackCircle, vbCrLf)

    ' ●と●の間の分配列を作成する
    ReDim Result_BlackCircle_Lines(UBound(BlackCircle_Lines))

    ' 最初と最後の文字列を指定することで、中間の文字列を抽出している
    For Extract_text_BlackCircle_i = LBound(BlackCircle_Lines) To UBound(BlackCircle_Lines) - 1
        Result_BlackCircle_Lines(Extract_text_BlackCircle_i) = ExtractTextBetweenTargets(Extract_text_BlackCircle_input, BlackCircle_Lines(Extract_text_BlackCircle_i), BlackCircle_Lines(Extract_text_BlackCircle_i + 1))
        ' Debug.Print ("======" & Result_BlackCircle_Lines(Extract_text_BlackCircle_i))
    Next Extract_text_BlackCircle_i

    ' Debug.Print ("段数：" & UBound(BlackCircle_Lines))
    Extract_text_BlackCircle = Result_BlackCircle_Lines
End Function

' ###########################################################################################

' ###########################################################################################
        ' ドット版
'
Private Function Extract_text_Dots(Extract_text_Dots_input As String)
    Dim Dots As String
    Dim Dots_Lines() As String
    Dim Result_Dots_Lines() As String
    Dim Extract_text_Dots_i As Long

    ' ●の列を抽出
    Dots = ExtractLines(Extract_text_Dots_input, "・")

    ' ●の配列を生成「絶対に存在しない●●●●●という文字列を加えている」
    Dots = Dots & vbCrLf & "・・・・・"

    ' 配列化している
    Dots_Lines = Split(Dots, vbCrLf)

    ' ●と●の間の分配列を作成する
    ReDim Result_Dots_Lines(UBound(Dots_Lines))

    ' 最初と最後の文字列を指定することで、中間の文字列を抽出している
    For Extract_text_Dots_i = LBound(Dots_Lines) To UBound(Dots_Lines) - 1
        Result_Dots_Lines(Extract_text_Dots_i) = ExtractTextBetweenTargets(Extract_text_Dots_input, Dots_Lines(Extract_text_Dots_i), Dots_Lines(Extract_text_Dots_i + 1))
        Extract_text_Dots_input = Replace(Extract_text_Dots_input, Result_Dots_Lines(Extract_text_Dots_i), "", 1, 1, 0)
        Debug.Print ("======" & Result_Dots_Lines(Extract_text_Dots_i))
    ' Next Extract_text_Dots_i

    Debug.Print ("段数：" & UBound(Dots_Lines))
    Extract_text_Dots = Result_Dots_Lines
End Function

' ###########################################################################################

