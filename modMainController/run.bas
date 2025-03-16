' パブリック変数の宣言 (モジュール4にまとめる)
Public LoadedText As String
Public processedText As String
Public ExtractedPatterns As Collection
Public ExtractedNumbers As Object

' メインの処理 (パブリック)
Public Sub RunAllModules()
    Call ClearImmediateWindow
    Call ReadFile
    Call ProcessText
End Sub

' ウィンドウのクリア
Private Sub ClearImmediateWindow()
    Debug.Print WorksheetFunction.Rept(vbLf, 200)
End Sub