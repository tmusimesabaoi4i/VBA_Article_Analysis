Sub GenerateReferencesFromSheet2()
    Dim ws2 As Worksheet, ws3 As Worksheet
    Dim lastRow2 As Long, lastRow3 As Long
    Dim i As Long, j As Long
    Dim bibNo As String, verNo As String
    Dim author As String, title As String, dateStr As String
    Dim rawDate As String, parts As Variant
    Dim url As String, urlParts As Variant
    Dim kk As String
    Dim todayStr As String
    Dim foundMatch As Boolean
    
    Set ws2 = ThisWorkbook.Sheets("Sheet2")
    Set ws3 = ThisWorkbook.Sheets("Sheet3")
    
    lastRow2 = ws2.Cells(ws2.Rows.Count, "C").End(xlUp).Row
    lastRow3 = ws3.Cells(ws3.Rows.Count, "A").End(xlUp).Row
    
    todayStr = Format(Date, "yyyy年mm月dd日")
    
    ' シート3の各行をループ
    For i = 2 To lastRow3
        bibNo = Trim(ws3.Cells(i, "A").Value)
        verNo = Trim(ws3.Cells(i, "B").Value)
        foundMatch = False
        
        ' シート2内で一致行を検索
        For j = 2 To lastRow2
            If Trim(ws2.Cells(j, "C").Value) = bibNo And Trim(ws2.Cells(j, "D").Value) = verNo Then
                foundMatch = True
                
                ' --- 著者・タイトル ---
                ws3.Cells(i, "C").Value = ws2.Cells(j, "G").Value
                ws3.Cells(i, "D").Value = ws2.Cells(j, "F").Value
                
                ' --- 日付（DD-MM-YYYY → YYYY/MM/DD） ---
                rawDate = Trim(ws2.Cells(j, "H").Value)
                If rawDate <> "" Then
                    parts = Split(rawDate, "-")
                    If UBound(parts) = 2 Then
                        dateStr = parts(2) & "/" & Format(parts(1), "00") & "/" & Format(parts(0), "00")
                        ws3.Cells(i, "E").Value = dateStr
                    End If
                End If
                
                ' --- URL抽出 ---
                On Error Resume Next
                url = ws2.Cells(j, "I").Hyperlinks(1).Address
                On Error GoTo 0
                ws3.Cells(i, "F").Value = url
                
                ' --- 取得日情報 ---
                If url <> "" Then
                    ws3.Cells(i, "G").Value = "[取得日 " & todayStr & "], 取得先 <" & url & ">"
                Else
                    ws3.Cells(i, "G").Value = "[取得日 " & todayStr & "], 取得先 <URLなし>"
                End If
                
                ' --- IEEE表記用 KK作成 ---
                If url <> "" Then
                    urlParts = Split(Mid(url, InStrRev(url, "/") + 1), "-")
                    If UBound(urlParts) >= 1 Then
                        kk = urlParts(0) & "-" & urlParts(1)
                    Else
                        kk = urlParts(0)
                    End If
                Else
                    kk = "N/A"
                End If
                
                ' --- IEEE文献表記作成 ---
                ws3.Cells(i, "H").Value = "IEEE802.11 " & kk & "/" & bibNo & "r" & verNo
                
                Exit For
            End If
        Next j
        
        If Not foundMatch Then
            ws3.Cells(i, "C").Value = "一致なし"
            ws3.Cells(i, "D").Value = ""
            ws3.Cells(i, "E").Value = ""
            ws3.Cells(i, "F").Value = ""
            ws3.Cells(i, "G").Value = ""
            ws3.Cells(i, "H").Value = ""
        End If
    Next i
    
    MsgBox "文献リファレンスの転記が完了しました。", vbInformation
End Sub