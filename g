Option Explicit

Private CancelFlag As Boolean

' ==== プログレスバー更新 ====
Public Sub UpdateProgress(ByVal current As Long, ByVal total As Long, ByVal statusText As String)
    With frmProgress
        Dim pct As Double
        pct = current / total
        If pct > 1 Then pct = 1
        .lblBar.Width = pct * .fraProgress.Width
        .lblPercent.Caption = Format(pct * 100, "0") & "%"
        .lblStatus.Caption = statusText
        DoEvents
    End With
End Sub

Public Sub CancelOperation()
    CancelFlag = True
End Sub


' ==== メイン処理 ====
Sub DownloadAndCombinePPTX_UI()
    Dim ws As Worksheet
    Dim baseFolder As String, mainFolder As String
    Dim fileURL As String, docNum As String, ver As String
    Dim row As Long, lastRow As Long
    Dim subFolder As String, pptxPath As String
    Dim combinedPPTX As String, htmlPath As String
    Dim appPPT As Object, pres As Object
    Dim combinePres As Object
    
    Set ws = ThisWorkbook.Sheets(1)
    baseFolder = Replace(Environ("USERPROFILE"), "C:", "E:") & "\Downloads\"
    mainFolder = baseFolder & ws.Range("A1").Value
    If Dir(mainFolder, vbDirectory) = "" Then MkDir mainFolder
    
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ' ==== フォーム起動 ====
    frmProgress.Show vbModeless
    CancelFlag = False
    UpdateProgress 0, lastRow, "初期化中..."
    
    Set appPPT = CreateObject("PowerPoint.Application")
    appPPT.Visible = True
    Set combinePres = appPPT.Presentations.Add
    
    Dim successCount As Long
    successCount = 0
    
    For row = 2 To lastRow
        If CancelFlag Then
            frmProgress.lblStatus.Caption = "❌ 中止されました"
            Exit For
        End If
        
        docNum = Trim(ws.Cells(row, "C").Value)
        ver = Trim(ws.Cells(row, "D").Value)
        fileURL = GetHyperlinkAddress(ws.Cells(row, "I"))
        
        If docNum <> "" And ver <> "" And fileURL <> "" Then
            subFolder = mainFolder & "\" & docNum & "_r" & ver
            If Dir(subFolder, vbDirectory) = "" Then MkDir subFolder
            
            pptxPath = subFolder & "\" & docNum & "_r" & ver & ".pptx"
            UpdateProgress row - 1, lastRow, "ダウンロード中: " & docNum
            
            DownloadFile_UI fileURL, pptxPath, "", 30, 5
            
            On Error Resume Next
            Set pres = appPPT.Presentations.Open(pptxPath, , , msoFalse)
            If Not pres Is Nothing Then
                pres.Slides.Range.Copy
                combinePres.Slides.Paste
                pres.Close
                successCount = successCount + 1
            End If
            On Error GoTo 0
        End If
    Next row
    
    combinedPPTX = mainFolder & "\combine_" & ws.Range("A1").Value & ".pptx"
    combinePres.SaveAs combinedPPTX
    htmlPath = mainFolder & "\combine_" & ws.Range("A1").Value & ".html"
    combinePres.SaveAs htmlPath, 12
    
    combinePres.Close
    appPPT.Quit
    
    UpdateProgress lastRow, lastRow, "✅ 完了: " & successCount & "件処理しました"
    MsgBox "完了しました！", vbInformation
    Unload frmProgress
End Sub


' ==== ダウンロード関数（UI更新対応） ====
Sub DownloadFile_UI(URL As String, SavePath As String, Optional Proxy As String = "", Optional Timeout As Long = 30, Optional Retry As Integer = 5)
    Dim xmlhttp As Object, adoStream As Object
    Dim attempt As Integer
    Dim success As Boolean
    
    success = False
    For attempt = 1 To Retry
        If CancelFlag Then Exit Sub
        
        UpdateProgress attempt, Retry, "通信中... (" & attempt & "/" & Retry & ")"
        
        Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
        On Error Resume Next
        xmlhttp.Open "GET", URL, False
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
            Exit For
        End If
        
        Application.Wait (Now + TimeValue("0:00:01"))
    Next attempt
    
    If Not success Then
        UpdateProgress Retry, Retry, "❌ 失敗: " & URL
    End If
End Sub


Function GetHyperlinkAddress(rng As Range) As String
    On Error Resume Next
    GetHyperlinkAddress = rng.Hyperlinks(1).Address
    On Error GoTo 0
End Function