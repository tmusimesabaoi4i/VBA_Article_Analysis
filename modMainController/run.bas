' パブリック変数の宣言 (モジュール4にまとめる)
Public LoadedText As String
Public processedText As String
Public ExtractedPatterns As Collection
Public ExtractedNumbers As Dictionary

' メインの処理 (パブリック)
Public Sub RunAllModules()
    Call ReadFile
End Sub
