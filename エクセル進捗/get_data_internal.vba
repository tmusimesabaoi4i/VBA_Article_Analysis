' 参照設定: VBAエディタ -> ツール -> 参照設定 で
'  - Microsoft HTML Object Library を ON にしておく（HTML解析用）
' （WinHttp は CreateObject で使うので参照不要）

' ---------- 汎用ヘルパードライバ ----------
Function SendHttpRequest(method As String, url As String, headers As Variant, Optional body As String = "") As String
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    http.Open UCase(method), url, False
    
    Dim i As Long
    If Not IsMissing(headers) Then
        For i = LBound(headers) To UBound(headers)
            ' headers は "Name: Value" の文字列配列を想定
            Dim parts() As String
            parts = Split(headers(i), ":", 2)
            If UBound(parts) >= 1 Then
                http.setRequestHeader Trim(parts(0)), Trim(parts(1))
            End If
        Next i
    End If
    
    If Len(body) > 0 Then
        http.Send body
    Else
        http.Send
    End If
    
    If http.Status >= 200 And http.Status < 300 Then
        SendHttpRequest = http.ResponseText
    Else
        Err.Raise vbObjectError + 1000, "SendHttpRequest", "HTTP Error: " & http.Status & " " & http.StatusText & vbCrLf & Left(http.ResponseText, 1000)
    End If
End Function

' ---------- cURL を手で分解して使う例（POST） ----------
Sub Example_FromCurl_POST()
    Dim url As String
    url = "https://internal.example.local/api/getReport" ' Copy as cURL から取った URL に置き換えて

    ' Copy as cURL にある -H '...' をこの形で配列に入れる
    Dim headers As Variant
    headers = Array( _
        "Accept: application/json, text/javascript, */*; q=0.01", _
        "Content-Type: application/json;charset=UTF-8", _
        "Cookie: sessionid=abcdef123456; other=xxx", _
        "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) ..." _
    )

    ' Copy as cURL の --data-raw '...' を body にコピー
    Dim body As String
    body = "{""param1"":""value1"",""param2"":123}" ' ここは cURL の data 部分をそのまま

    Dim resp As String
    On Error GoTo ErrHandler
    resp = SendHttpRequest("POST", url, headers, body)
    Debug.Print resp  ' イミディエイトに出す。長ければファイルに保存してもOK

    ' JSONならここで Parse する（外部ライブラリ不要ならテキスト処理）
    Exit Sub

ErrHandler:
    MsgBox "エラー: " & Err.Description, vbExclamation
End Sub

' ---------- GET の例 ----------
Sub Example_FromCurl_GET()
    Dim url As String
    url = "https://internal.example.local/data?x=1"

    Dim headers As Variant
    headers = Array( _
        "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8", _
        "Cookie: sessionid=abcdef123456; other=xxx", _
        "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) ..." _
    )

    Dim resp As String
    resp = SendHttpRequest("GET", url, headers)
    ' HTML を解析
    Dim doc As New MSHTML.HTMLDocument
    doc.body.innerHTML = resp

    ' 例: id="result-table" の中身を取る
    Dim tbl As MSHTML.IHTMLElement
    Set tbl = doc.getElementById("result-table")
    If Not tbl Is Nothing Then
        Debug.Print tbl.innerText
    Else
        Debug.Print "要素が見つかりません"
    End If
End Sub