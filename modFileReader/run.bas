' メインの処理 (パブリック)
    '
    ' LoadedText:読み込んだテキスト(jis)
    ' processedText:半角、改行削除済テキスト(jis)

Private filePath As String

Public Sub ReadFile()
    Call getFilePathFromCellA1
    Call readFileAsJIS
    Call removeNewlinesAndConvertToHalfwidth
End Sub

' ファイルの読み込み (プライベート)
Private Sub getFilePathFromCellA1()
    Dim filePath_tmp As String

    filePath_tmp = Worksheets(1).Cells(1, 1).Value

    ' 不要な文字の削除
    filePath_tmp = Replace(filePath_tmp, " ", "")     '半角スペース削除
    filePath_tmp = Replace(filePath_tmp, "　", "")    '全角スペース削除
    filePath_tmp = Replace(filePath_tmp, vbTab, "")   'タブ削除
    filePath_tmp = Replace(filePath_tmp, vbVerticalTab, "")   'タブ削除
    filePath_tmp = Replace(filePath_tmp, vbCrLf, "")    'セル内改行削除
    filePath_tmp = Replace(filePath_tmp, vbCr, "")    'セル内改行削除
    filePath_tmp = Replace(filePath_tmp, vbLf, "")    'セル内改行削除
    filePath_tmp = Replace(filePath_tmp, vbNewLine, "")    'セル内改行削除

    filePath = ThisWorkbook.Path & "\" & filePath_tmp

    ' 出力
    Debug.Print ("ファイル名：" & filePath_tmp)
    Debug.Print ("ファイルパス：" & filePath)
End Sub

' ファイルの読み込み (プライベート)
Private Sub readFileAsJIS()
    Dim LoadedText_tmp As String

    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile filePath
        LoadedText_tmp = .ReadText
        .Close
    End With
    
    LoadedText = LoadedText_tmp

    ' 出力
    ' Debug.Print ("入力文字(処理前)：" & LoadedText)
End Sub

' ファイルの読み込み (プライベート)
Private Sub removeNewlinesAndConvertToHalfwidth()
    Dim processedText_tmp As String

    processedText_tmp = LoadedText

    ' 不要な文字の削除
    processedText_tmp = Replace(processedText_tmp, " ", "")     '半角スペース削除
    processedText_tmp = Replace(processedText_tmp, "　", "")    '全角スペース削除
    processedText_tmp = Replace(processedText_tmp, vbTab, "")   'タブ削除
    processedText_tmp = Replace(processedText_tmp, vbVerticalTab, "")   'タブ削除
    processedText_tmp = Replace(processedText_tmp, vbCrLf, "")    'セル内改行削除
    processedText_tmp = Replace(processedText_tmp, vbCr, "")    'セル内改行削除
    processedText_tmp = Replace(processedText_tmp, vbLf, "")    'セル内改行削除
    processedText_tmp = Replace(processedText_tmp, vbNewLine, "")    'セル内改行削除

    processedText = processedText_tmp

    ' 出力
    ' Debug.Print ("入力文字：" & processedText)
End Sub


