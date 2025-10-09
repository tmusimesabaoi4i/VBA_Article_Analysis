Option Explicit

' ===== メイン処理 =====
Sub DownloadAndCombinePPTX()
    Dim ws As Worksheet
    Dim baseFolder As String, mainFolder As String
    Dim fileURL As String, docNum As String, ver As String
    Dim row As Long, lastRow As Long
    Dim subFolder As String, pptxPath As String
    Dim combinedPPTX As String, htmlPath As String
    Dim appPPT As Object, pres As Object
    Dim combinePres As Object
    
    Set ws = ThisWorkbook.Sheets(1)
    
    ' --- ✅ USERPROFILE のドライブ文字を E に強制 ---
    baseFolder = Replace(Environ("USERPROFILE"), "C:", "E:") & "\Downloads\"
    
    mainFolder = baseFolder & ws.Range("A1").Value
    If Dir(mainFolder, vbDirectory) = "" Then MkDir mainFolder
    
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    Debug.Print "===== ダウンロード開始 ====="
    
    ' PowerPoint起動
    Set appPPT = CreateObject("PowerPoint.Application")
    appPPT.Visible = True
    Set combinePres = appPPT.Presentations.Add

    For row = 2 To lastRow
        docNum = Trim(ws.Cells(row, "C").Value)
        ver = Trim(ws.Cells(row, "D").Value)
        fileURL = GetHyperlinkAddress(ws.Cells(row, "I"))
        
        If docNum <> "" And ver <> "" And fileURL <> "" Then
            subFolder = mainFolder & "\" & docNum & "_r" & ver
            If Dir(subFolder, vbDirectory) = "" Then MkDir subFolder
            
            pptxPath = subFolder & "\" & docNum & "_r" & ver & ".pptx"
            Debug.Print "▶ ダウンロード中: " & fileURL
            DownloadFile fileURL, pptxPath, "", 30, 10
            
            ' PPTXを結合
            On Error Resume Next
            Set pres = appPPT.Presentations.Open(pptxPath, , , msoFalse)
            If Not pres Is Nothing Then
                pres.Slides.Range.Copy
                combinePres.Slides.Paste
                pres.Close
            End If
            On Error GoTo 0
        End If
    Next row

    ' 結合ファイルを保存
    combinedPPTX = mainFolder & "\combine_" & ws.Range("A1").Value & ".pptx"
    combinePres.SaveAs combinedPPTX
    Debug.Print "✅ 結合完了: " & combinedPPTX

    ' HTMLに変換
    htmlPath = mainFolder & "\combine_" & ws.Range("A1").Value & ".html"
    combinePres.SaveAs htmlPath, 12 'ppSaveAsHTML = 12
    Debug.Print "✅ HTML変換完了: " & htmlPath

    combinePres.Close
    appPPT.Quit
    Debug.Print "===== 完了しました ====="
End Sub


' ===== ハイパーリンク取得 =====
Function GetHyperlinkAddress(rng As Range) As String
    On Error Resume Next
    GetHyperlinkAddress = rng.Hyperlinks(1).Address
    On Error GoTo 0
End Function


' ===== ファイルダウンロード関数 =====
Function DownloadFile(URL As String, SavePath As String, Optional Proxy As String = "", Optional Timeout As Long = 30, Optional Retry As Integer = 10)
    Dim xmlhttp As Object
    Dim adoStream As Object
    Dim attempt As Integer
    Dim success As Boolean
    Dim startTime As Double
    
    success = False
    For attempt = 1 To Retry
        Debug.Print "通信中... [" & attempt & "回目]"
        Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
        startTime = Timer
        On Error Resume Next
        xmlhttp.Open "GET", URL, False
        If Proxy <> "" Then
            xmlhttp.setProxy 2, Proxy, ""
        End If
        xmlhttp.setTimeouts 0, Timeout * 1000, Timeout * 1000, Timeout * 1000
        xmlhttp.Send
        On Error GoTo 0
        
        If xmlhttp.Status = 200 Then
            Set adoStream = CreateObject("ADODB.Stream")
            adoStream.Type = 1
            adoStream.Open
            adoStream.Write xmlhttp.responseBody
            adoStream.SaveToFile SavePath, 2
            adoStream.Close
            success = True
            Debug.Print "✅ ダウンロード成功 (" & Round(Timer - startTime, 1) & "秒)"
            Exit For
        Else
            Debug.Print "⚠️ リトライ中... (HTTP " & xmlhttp.Status & ")"
            DoEvents
            Application.Wait (Now + TimeValue("0:00:02"))
        End If
    Next attempt
    
    If Not success Then
        Debug.Print "❌ ダウンロード失敗: " & URL
    End If
End Function