Option Explicit

Sub FilterAndSortByKeyword()
    Dim wsSrc As Worksheet, wsDst As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long, dstRow As Long
    Dim titleCol As Long
    Dim keyword As String
    Dim matchFlag As Boolean
    Dim cellTitle As String
    
    ' ===== シート設定 =====
    Set wsSrc = ThisWorkbook.Sheets("SheetF")   ' ★シート名は必要に応じて変更
    Set wsDst = ThisWorkbook.Sheets("Sheet2")
    wsDst.Cells.Clear
    
    ' ===== 行・列情報 =====
    titleCol = 6   ' F列 = タイトル列
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, titleCol).End(xlUp).Row
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    
    ' ===== キーワード取得（1行目のみ）=====
    Dim keywords() As String
    Dim kwCount As Long
    kwCount = lastCol
    ReDim keywords(1 To kwCount)
    For j = 1 To kwCount
        ' --- 前後空白・改行を除去 ---
        keywords(j) = CleanKeyword(wsSrc.Cells(1, j).Value)
    Next j
    
    ' ===== 見出し行コピー =====
    wsSrc.Rows(1).Copy wsDst.Rows(1)
    dstRow = 2
    
    ' ===== タイトル抽出ループ（2行目～）=====
    For i = 2 To lastRow
        cellTitle = LCase(Trim(wsSrc.Cells(i, titleCol).Value)) ' 小文字化で統一
        matchFlag = False
        
        For j = 1 To kwCount
            If keywords(j) <> "" Then
                If InStr(1, cellTitle, LCase(keywords(j)), vbTextCompare) > 0 Then
                    matchFlag = True
                    Exit For
                End If
            End If
        Next j
        
        If matchFlag Then
            wsSrc.Rows(i).Copy wsDst.Rows(dstRow)
            dstRow = dstRow + 1
        End If
    Next i
    
    ' ===== 並べ替え =====
    Dim sortRange As Range
    Dim lastRowDst As Long
    lastRowDst = wsDst.Cells(wsDst.Rows.Count, "C").End(xlUp).Row
    
    If lastRowDst < 2 Then
        MsgBox "該当するタイトルが見つかりませんでした。", vbInformation
        Exit Sub
    End If
    
    Set sortRange = wsDst.Range("A1").CurrentRegion
    
    sortRange.Sort Key1:=wsDst.Range("C2"), Order1:=xlAscending, _
                   Key2:=wsDst.Range("D2"), Order2:=xlAscending, _
                   Key3:=wsDst.Range("H2"), Order3:=xlAscending, _
                   Header:=xlYes
    
    MsgBox "抽出と整列が完了しました。", vbInformation
End Sub


' ===== キーワード整形（空白・改行を除去）=====
Private Function CleanKeyword(ByVal s As String) As String
    s = Trim(s)
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, "　", "") ' 全角スペース除去
    CleanKeyword = s
End Function