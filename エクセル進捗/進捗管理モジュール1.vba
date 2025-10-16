Option Explicit

'====================================
' メイン実行関数（手動実行用）
'====================================
Sub RunProgressApp()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' データ更新と再計算
    UpdateSheet3ToSheet2
    UpdateMonthlyProgress
    
    ' UI整形
    SetupUI

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "✅ 進捗管理アプリの更新が完了しました。", vbInformation
End Sub

'====================================
' Sheet3変更時に自動トリガー
'====================================
Private Sub Worksheet_Change(ByVal Target As Range)
    ' このイベントは Sheet3 に貼り付ける（ThisWorkbookではなく、Sheet3モジュール）
    If Not Intersect(Target, Me.Range("A:C")) Is Nothing Then
        Application.EnableEvents = False
        RunProgressApp
        Application.EnableEvents = True
    End If
End Sub

'====================================
' Sheet3→Sheet2へ転記処理
'====================================
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

'====================================
' Sheet2→Sheet1の月次集計
'====================================
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
        ' 進捗率 = E / D
        If wsMonth.Cells(i, "D").Value <> 0 Then
            wsMonth.Cells(i, "F").Value = wsMonth.Cells(i, "E").Value / wsMonth.Cells(i, "D").Value
        End If

        ' 平日ベース進捗比較
        wsMonth.Cells(i, "G").Value = GetProgressDeviation(y, m, wsMonth.Cells(i, "F").Value)
    Next i

    ' データバーを設定
    Set rng = wsMonth.Range("F2:F" & lastRowMonth)
    Call ApplyDataBar(rng)
End Sub

'====================================
' データバー適用
'====================================
Sub ApplyDataBar(rng As Range)
    Dim cf As FormatCondition
    rng.FormatConditions.Delete
    Set cf = rng.FormatConditions.AddDatabar()
    cf.MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
    cf.MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=1
    cf.BarColor.Color = RGB(91, 155, 213)
End Sub

'====================================
' 遅れ指標計算
'====================================
Function GetProgressDeviation(y As Long, m As Long, currentProgress As Double) As Double
    Dim firstDate As Date, lastDate As Date
    Dim today As Date, weekdaysTotal As Long, weekdaysPassed As Long
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

    If weekdaysTotal = 0 Then
        GetProgressDeviation = 0
    Else
        Dim expectedProgress As Double
        expectedProgress = weekdaysPassed / weekdaysTotal
        GetProgressDeviation = currentProgress - expectedProgress
    End If
End Function

'====================================
' Sheet2の遅延案件ハイライト
'====================================
Sub HighlightOverdueTasks()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim dueDate As Variant, doneDate As Variant

    Set ws = ThisWorkbook.Sheets("Sheet2")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ws.Range("B2:E" & lastRow).Interior.ColorIndex = xlNone

    For i = 2 To lastRow
        dueDate = ws.Cells(i, "C").Value
        doneDate = ws.Cells(i, "E").Value
        If IsDate(dueDate) And doneDate = "" Then
            If dueDate < Date Then
                ws.Range("B" & i & ":E" & i).Interior.Color = RGB(255, 199, 206)
            End If
        End If
    Next i
End Sub

'====================================
' --- UI整形系 ---
'====================================
Sub SetupUI()
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    StyleHeaders
    ApplyBorders
    HighlightOverdueTasks
    ColorByOverdueDegree
    ColorByDeviation

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub StyleHeaders()
    Dim ws As Worksheet
    Dim headerColor As Long
    Dim fontColor As Long

    headerColor = RGB(91, 155, 213)
    fontColor = RGB(255, 255, 255)

    For Each ws In ThisWorkbook.Worksheets
        With ws.Rows(1)
            .Font.Bold = True
            .Interior.Color = headerColor
            .Font.Color = fontColor
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .RowHeight = 22
        End With
    Next ws
End Sub

Sub ApplyBorders()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim rng As Range

    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        lastRow = ws.Cells.Find("*", , , , xlByRows, xlPrevious).Row
        lastCol = ws.Cells.Find("*", , , , xlByColumns, xlPrevious).Column
        On Error GoTo 0

        If lastRow > 1 And lastCol > 1 Then
            Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
            With rng.Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(200, 200, 200)
            End With
        End If
    Next ws
End Sub

Sub ColorByOverdueDegree()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim dueDate As Variant, doneDate As Variant
    Dim diffDays As Long
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
                    Case 1 To 3
                        colorCode = RGB(255, 235, 156)
                    Case 4 To 7
                        colorCode = RGB(255, 192, 0)
                    Case Is > 7
                        colorCode = RGB(255, 99, 71)
                    Case Else
                        colorCode = RGB(255, 255, 255)
                End Select
                ws.Range("B" & i & ":E" & i).Interior.Color = colorCode
            End If
        End If
    Next i
End Sub

Sub ColorByDeviation()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim deviation As Double
    Dim colorCode As Long

    Set ws = ThisWorkbook.Sheets("Sheet1")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    For i = 2 To lastRow
        deviation = ws.Cells(i, "G").Value
        Select Case True
            Case deviation < -0.15
                colorCode = RGB(255, 99, 71)
            Case deviation < -0.05
                colorCode = RGB(255, 192, 0)
            Case deviation < 0.05
                colorCode = RGB(198, 239, 206)
            Case Else
                colorCode = RGB(142, 209, 123)
        End Select
        ws.Cells(i, "G").Interior.Color = colorCode
    Next i
End Sub