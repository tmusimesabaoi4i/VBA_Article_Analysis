Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("A:C")) Is Nothing Then
        Application.EnableEvents = False
        RunProgressApp
        Application.EnableEvents = True
    End If
End Sub