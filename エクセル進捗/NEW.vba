Option Explicit

'===========================================================
' 📦 メイン実行関数
'===========================================================
Sub RunProgressApp()
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' データ更新処理
    UpdateSheet3ToSheet2
    UpdateMonthlyProgress

    ' UI整形
    SetupUI

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "✅ 進捗管理アプリの更新が完了しました。", vbInformation
End Sub


'===========================================================
' 🔄 Sheet3 → Sheet2 転記処理
'===========================================================
Sub UpdateSheet3ToSheet2()
    Dim wsData As Worksheet, wsTask As Worksheet
    Dim lastRowData As Long, lastRowTask As Long
    Dim i As Long, j As Long
    Dim idData As String, idTask As String

    Set wsData = ThisWorkbook.Sheets("Sheet3")
    Set wsTask = ThisWorkbook.Sheets("Sheet2")

    lastRowData = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    lastRowTask = wsTask.Cells(wsTask.Rows.Count, "B").End(xlUp).Row

    For i = 2 To lastRowData
        idData = Trim(wsData.Cells(i, "A").Value)
        If idData <> "" Then
            For j = 2 To lastRowTask
                idTask = Trim(wsTask.Cells(j, "B").Value)
                If idTask = idData Then
                    wsTask.Cells(j, "E").Value = wsData.Cells(i, "B").Value '達成日
                    wsTask.Cells(j, "D").Value = wsData.Cells(i, "C").Value '達成値
                    Exit For
                End If
            Next j
        End If
    Next i
End Sub


'===========================================================
' 📊 月次進捗更新 (Sheet2 → Sheet1)
'===========================================================
Sub UpdateMonthlyProgress()
    Dim wsMonth As Worksheet, wsTask As Worksheet
    Dim lastRowMonth As Long, lastRowTask As Long
    Dim i As Long, j As Long
    Dim y As Long, m As Long
    Dim sumVal As Double
    Dim targetY As Long, targetM As Long
    Dim rng As Range

    Set wsMonth = ThisWorkbook.Sheets("Sheet1")
    Set wsTask = ThisWorkbook.Sheets("Sheet2")

    lastRowMonth = wsMonth.Cells(wsMonth.Rows.Count, "B").End(xlUp).Row
    lastRowTask = wsTask.Cells(wsTask.Rows.Count, "B").End(xlUp).Row

    For i = 2 To lastRowMonth
        y = wsMonth.Cells(i, "B").Value
        m = wsMonth.Cells(i, "C").Value
        sumVal = 0

        For j = 2 To lastRowTask
            If Not IsEmpty(wsTask.Cells(j, "E").Value) Then
                targetY = Year(wsTask.Cells(j, "E").Value)
                targetM = Month(wsTask.Cells(j, "E").Value)
                If targetY = y And targetM = m Then
                    sumVal = sumVal + wsTask.Cells(j, "D").Value
                End If
            End If
        Next j

        wsMonth.Cells(i, "E").Value = sumVal
        If wsMonth.Cells(i, "D").Value <> 0 Then
            wsMonth.Cells(i, "F").Value = wsMonth.Cells(i, "E").Value / wsMonth.Cells(i, "D").Value
        End If
        wsMonth.Cells(i, "G").Value = GetProgressDeviation(y, m, wsMonth.Cells(i, "F").Value)
    Next i

    ' データバー
    Set rng = wsMonth.Range("F2:F" & lastRowMonth)
    ApplyDataBar rng
End Sub


'===========================================================
' 📈 データバー適用（安全版）
'===========================================================
'====================================
' データバー適用（安全版：自動フォールバック付き）
'====================================
Sub ApplyDataBar(rng As Range)
    Dim lastRow As Long
    On Error Resume Next
    lastRow = rng.Parent.Cells(rng.Parent.Rows.Count, rng.Column).End(xlUp).Row
    On Error GoTo 0
    If lastRow < 2 Then Exit Sub   ' データなし

    ' 範囲が逆転していないか（F2:F1 など）
    If rng.Row > rng.Rows(rng.Rows.Count).Row Then Exit Sub

    ' 既存CF削除
    rng.FormatConditions.Delete

    ' データバー対応判定：バージョン & ファイル形式
    If SupportsDataBars() Then
        Dim db As DataBar
        On Error Resume Next
        Set db = rng.FormatConditions.AddDatabar
        If Err.Number <> 0 Or db Is Nothing Then
            Err.Clear
            On Error GoTo 0
            ' フォールバック
            ApplyColorScaleFallback rng
            Exit Sub
        End If
        On Error GoTo 0

        With db
            .MinPoint.Modify Type:=xlConditionValueNumber, Value:=0
            .MaxPoint.Modify Type:=xlConditionValueNumber, Value:=1
            .BarFillType = xlDataBarFillSolid
            .BarColor.Color = RGB(91, 155, 213)
            .AxisPosition = xlDataBarAxisAutomatic
            .ShowValue = True
        End With

        ' 進捗率用の表示形式（任意）
        On Error Resume Next
        rng.NumberFormatLocal = "0%"
        On Error GoTo 0
    Else
        ' フォールバック
        ApplyColorScaleFallback rng
    End If
End Sub

'====================================
' データバー対応可否判定
' - Excel 2007以降 & xlsx/xlsm 等の新形式なら True
'====================================
Function SupportsDataBars() As Boolean
    Dim ver As Double
    Dim ff As Long
    On Error Resume Next
    ver = CDbl(Application.Version)        ' 12=2007, 14=2010, ...
    ff = ThisWorkbook.FileFormat
    On Error GoTo 0

    If ver >= 12 Then
        ' 互換モード（xls=xlExcel8）は不可
        If ff <> xlExcel8 And ff <> xlExcel4Workbook Then
            SupportsDataBars = True
            Exit Function
        End If
    End If
    SupportsDataBars = False
End Function

'====================================
' フォールバック1：2色カラー スケール（下限→上限）
' データバー不可時の代替（Excel 2007 以降）
'====================================
Sub ApplyColorScaleFallback(rng As Range)
    On Error GoTo HardFallback

    Dim cs As ColorScale
    ' 既存CFを消してから実施（冪等性確保）
    rng.FormatConditions.Delete

    ' 2色スケール：0（赤）→ 1（緑）
    Set cs = rng.FormatConditions.AddColorScale(ColorScaleType:=2)
    With cs.ColorScaleCriteria(1)
        .Type = xlConditionValueNumber
        .Value = 0
        .FormatColor.Color = RGB(255, 99, 71)      ' 赤
    End With
    With cs.ColorScaleCriteria(2)
        .Type = xlConditionValueNumber
        .Value = 1
        .FormatColor.Color = RGB(142, 209, 123)    ' 緑
    End With

    On Error Resume Next
    rng.NumberFormatLocal = "0%"
    On Error GoTo 0
    Exit Sub

HardFallback:
    ' 最終フォールバック：しきい値で単色塗り分け（どの環境でも動く）
    ApplyThresholdFill rng
End Sub

'====================================
' フォールバック2：しきい値の単色塗り（最終手段）
' - <0.3 赤、<0.7 黄、>=0.7 緑
'====================================
Sub ApplyThresholdFill(rng As Range)
    Dim c As Range, v As Variant
    For Each c In rng.Cells
        v = c.Value
        If IsNumeric(v) Then
            If v < 0.3 Then
                c.Interior.Color = RGB(255, 99, 71)      ' 赤
            ElseIf v < 0.7 Then
                c.Interior.Color = RGB(255, 192, 0)      ' 黄
            Else
                c.Interior.Color = RGB(142, 209, 123)    ' 緑
            End If
        Else
            c.Interior.ColorIndex = xlNone
        End If
    Next c
End Sub


'===========================================================
' 🎨 単色しきい値（データバー代替）
'===========================================================
Sub ApplyThresholdFill(rng As Range)
    Dim c As Range, v As Variant
    For Each c In rng.Cells
        v = c.Value
        If IsNumeric(v) Then
            If v < 0.3 Then
                c.Interior.Color = RGB(255, 99, 71)
            ElseIf v < 0.7 Then
                c.Interior.Color = RGB(255, 192, 0)
            Else
                c.Interior.Color = RGB(142, 209, 123)
            End If
        Else
            c.Interior.ColorIndex = xlNone
        End If
    Next c
End Sub


'===========================================================
' ⏱ 平日ベース遅れ指標
'===========================================================
Function GetProgressDeviation(y As Long, m As Long, currentProgress As Double) As Double
    Dim firstDate As Date, lastDate As Date, today As Date
    Dim weekdaysTotal As Long, weekdaysPassed As Long
    Dim d As Date
    firstDate = DateSerial(y, m, 1)
    lastDate = DateSerial(y, m + 1, 0)
    today = Date
    For d = firstDate To lastDate
        If Weekday(d, vbMonday) <= 5 Then
            weekdaysTotal = weekdaysTotal + 1
            If d <= today Then weekdaysPassed = weekdaysPassed + 1
        End If
    Next d
    If weekdaysTotal = 0 Then Exit Function
    GetProgressDeviation = currentProgress - (weekdaysPassed / weekdaysTotal)
End Function


'===========================================================
' ⚠️ 遅延案件ハイライト
'===========================================================
Sub HighlightOverdueTasks()
    Dim ws As Worksheet, i As Long, lastRow As Long
    Dim dueDate As Variant, doneDate As Variant
    Set ws = ThisWorkbook.Sheets("Sheet2")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    ws.Range("B2:E" & lastRow).Interior.ColorIndex = xlNone
    For i = 2 To lastRow
        dueDate = ws.Cells(i, "C").Value
        doneDate = ws.Cells(i, "E").Value
        If IsDate(dueDate) And doneDate = "" And dueDate < Date Then
            ws.Range("B" & i & ":E" & i).Interior.Color = RGB(255, 199, 206)
        End If
    Next i
End Sub


'===========================================================
' 🎨 UI整形（完全版）
'===========================================================
Sub SetupUI()
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ApplyGlobalFont
    WriteHeaders
    StyleHeaders
    ApplyHeaderGradient
    ApplyBorders
    ColorizeDataRows
    HighlightOverdueTasks
    ColorByOverdueDegree
    ColorByDeviation

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "🎨 UI整形完了（フォント・配色・ヘッダ含む）", vbInformation
End Sub


'===========================================================
' ✍️ 各シートのヘッダタイトル
'===========================================================
Sub WriteHeaders()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Set ws1 = ThisWorkbook.Sheets("Sheet1")
    Set ws2 = ThisWorkbook.Sheets("Sheet2")

    ' Sheet1
    With ws1
        .Cells(1, "B").Value = "年"
        .Cells(1, "C").Value = "月"
        .Cells(1, "D").Value = "目標値"
        .Cells(1, "E").Value = "達成値"
        .Cells(1, "F").Value = "進捗率"
        .Cells(1, "G").Value = "遅れ指標"
        .Columns("B:G").AutoFit
    End With

    ' Sheet2
    With ws2
        .Cells(1, "B").Value = "案件番号"
        .Cells(1, "C").Value = "達成予定日"
        .Cells(1, "D").Value = "獲得達成値"
        .Cells(1, "E").Value = "達成日"
        .Columns("B:E").AutoFit
    End With
End Sub


'===========================================================
' 🌐 全体フォント統一
'===========================================================
Sub ApplyGlobalFont()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Cells.Font.Name = "メイリオ"
        ws.Cells.Font.Size = 11
    Next ws
End Sub


'===========================================================
' 🟦 ヘッダ装飾
'===========================================================
Sub StyleHeaders()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        With ws.Rows(1)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(91, 155, 213)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .RowHeight = 22
        End With
    Next ws
End Sub


'===========================================================
' 💎 ヘッダグラデーション
'===========================================================
Sub ApplyHeaderGradient()
    Dim ws As Worksheet, r As Range
    For Each ws In ThisWorkbook.Worksheets
        Set r = ws.Rows(1)
        If Application.WorksheetFunction.CountA(r) > 0 Then
            r.Interior.Pattern = xlPatternLinearGradient
            r.Interior.Gradient.Degree = 90
            With r.Interior.Gradient.ColorStops
                .Clear
                .Add(0).Color = RGB(91, 155, 213)
                .Add(1).Color = RGB(142, 180, 227)
            End With
        End If
    Next ws
End Sub


'===========================================================
' 🔲 枠線
'===========================================================
Sub ApplyBorders()
    Dim ws As Worksheet, lastRow As Long, lastCol As Long, rng As Range
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        lastRow = ws.Cells.Find("*", , , , xlByRows, xlPrevious).Row
        lastCol = ws.Cells.Find("*", , , , xlByColumns, xlPrevious).Column
        On Error GoTo 0
        If lastRow > 1 Then
            Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
            rng.Borders.LineStyle = xlContinuous
            rng.Borders.Weight = xlThin
            rng.Borders.Color = RGB(200, 200, 200)
        End If
    Next ws
End Sub


'===========================================================
' 🟦 交互色（縞模様）
'===========================================================
Sub ColorizeDataRows()
    Dim ws As Worksheet, lastRow As Long, lastCol As Long, i As Long
    Dim lightColor As Long, darkColor As Long
    lightColor = RGB(242, 242, 242)
    darkColor = RGB(217, 225, 242)

    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        lastRow = ws.Cells.Find("*", , , , xlByRows, xlPrevious).Row
        lastCol = ws.Cells.Find("*", , , , xlByColumns, xlPrevious).Column
        On Error GoTo 0
        If lastRow > 1 Then
            For i = 2 To lastRow
                If Application.WorksheetFunction.CountA(ws.Rows(i)) > 0 Then
                    ws.Rows(i).Interior.Color = IIf(i Mod 2 = 0, lightColor, darkColor)
                End If
            Next i
        End If
    Next ws
End Sub


'===========================================================
' ⚠️ 遅延度による色分け
'===========================================================
Sub ColorByOverdueDegree()
    Dim ws As Worksheet, i As Long, lastRow As Long
    Dim diffDays As Long, dueDate As Variant, doneDate As Variant
    Dim colorCode As Long
    Set ws = ThisWorkbook.Sheets("Sheet2")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    For i = 2 To lastRow
        dueDate = ws.Cells(i, "C").Value
        doneDate = ws.Cells(i, "E").Value
        If IsDate(dueDate) And doneDate = "" Then
            diffDays = Date - dueDate
            If diffDays > 0 Then
                Select Case diffDays
                    Case 1 To 3: colorCode = RGB(255, 235, 156)
                    Case 4 To 7: colorCode = RGB(255, 192, 0)
                    Case Is > 7: colorCode = RGB(255, 99, 71)
                End Select
                ws.Range("B" & i & ":E" & i).Interior.Color = colorCode
            End If
        End If
    Next i
End Sub


'===========================================================
' 📉 遅れ指標による色分け
'===========================================================
Sub ColorByDeviation()
    Dim ws As Worksheet, i As Long, lastRow As Long
    Dim dev As Double, colorCode As Long
    Set ws = ThisWorkbook.Sheets("Sheet1")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    For i = 2 To lastRow
        dev = ws.Cells(i, "G").Value
        Select Case True
            Case dev < -0.15: colorCode = RGB(255, 99, 71)
            Case dev < -0.05: colorCode = RGB(255, 192, 0)
            Case dev < 0.05: colorCode = RGB(198, 239, 206)
            Case Else: colorCode = RGB(142, 209, 123)
        End Select
        ws.Cells(i, "G").Interior.Color = colorCode
    Next i
End Sub