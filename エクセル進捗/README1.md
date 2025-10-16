# WinHttp + MSHTML を使った `<tr>` テーブル抽出サンプル（.md）

このサンプルでは、前述の「Copy as cURL」で取得した HTML ページに  
表（`<table>` ～ `<tr>` ～ `<td>`）が存在する場合、  
**表のヘッダー行 (`<th>`) および各データ行 (`<td>`) を Excel シートにそのまま転記** します。  
書式は無視して構いません。  

---

## 概要
- WinHttp で指定 URL の HTML を取得  
- MSHTML で DOM としてパース  
- `<table>` → `<tr>` → `<th>/<td>` の階層を順に解析  
- Excel シートの A1 から順に値を書き出し  

---

## VBA コード（テーブル抽出）

```vb
' ===============================================
' HTML の <table> を解析して Excel に転記するサンプル
' ===============================================

Option Explicit

' ▼ MSHTML を使うので参照設定：
' [VBE] → [ツール] → [参照設定] → "Microsoft HTML Object Library" にチェック

Sub Example_ExtractTableToExcel()
    On Error GoTo EH

    Dim url As String
    url = "https://internal.example.local/report" ' ← Copy as cURL の URL に差し替え

    ' 必要に応じてヘッダを設定（最低限 Accept と User-Agent）
    Dim headers As Variant
    headers = Array( _
        "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8", _
        "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64)" _
    )

    Dim htmlText As String
    htmlText = SendHttpGet(url, headers)

    ' HTML を DOM としてロード
    Dim doc As New MSHTML.HTMLDocument
    doc.Open
    doc.Write htmlText
    doc.Close

    ' テーブルを全取得（最初の1個でも複数でも可）
    Dim tbls As MSHTML.IHTMLElementCollection
    Set tbls = doc.getElementsByTagName("table")

    If tbls.Length = 0 Then
        MsgBox "テーブル要素が見つかりませんでした。", vbExclamation
        Exit Sub
    End If

    ' 出力先
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    ws.Cells.Clear

    Dim t As Long, r As Long, c As Long
    Dim trEls As MSHTML.IHTMLElementCollection
    Dim trEl As MSHTML.IHTMLElement
    Dim cellEls As MSHTML.IHTMLElementCollection
    Dim cellEl As MSHTML.IHTMLElement
    Dim rowIndex As Long, colIndex As Long

    rowIndex = 1

    ' 各テーブルを順に処理
    For t = 0 To tbls.Length - 1
        Set trEls = tbls.Item(t).getElementsByTagName("tr")

        For Each trEl In trEls
            Set cellEls = trEl.getElementsByTagName("th")
            If cellEls.Length = 0 Then
                Set cellEls = trEl.getElementsByTagName("td")
            End If

            colIndex = 1
            For Each cellEl In cellEls
                ws.Cells(rowIndex, colIndex).Value = Trim(cellEl.innerText)
                colIndex = colIndex + 1
            Next cellEl

            rowIndex = rowIndex + 1
        Next trEl

        ' テーブルが複数ある場合は1行空ける
        rowIndex = rowIndex + 1
    Next t

    ws.Columns.AutoFit
    MsgBox "完了：テーブルをシートに貼り付けました。", vbInformation
    Exit Sub

EH:
    MsgBox "エラー: " & Err.Description, vbExclamation
End Sub


' ===============================================
' GET でHTML取得（WinHttp使用）
' ===============================================
Private Function SendHttpGet(ByVal url As String, ByVal headers As Variant) As String
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "GET", url, False
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        Dim parts() As String
        parts = Split(headers(i), ":", 2)
        If UBound(parts) >= 1 Then
            http.setRequestHeader Trim(parts(0)), Trim(parts(1))
        End If
    Next i

    http.Send

    If http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + 2000, "SendHttpGet", _
                  "HTTP Error: " & http.Status & " " & http.StatusText
    End If

    SendHttpGet = http.ResponseText
End Function
```

---

## 実行結果イメージ

| A列 | B列 | C列 |  
|:----|:----|:----|  
| Header1 | Header2 | Header3 |  
| Data1 | Data2 | Data3 |  
| Data4 | Data5 | Data6 |  

> 💡 複数テーブルがあった場合は、各テーブルの間に **1行の空白行** を自動で挿入します。

---

## カスタマイズポイント

- **特定のテーブルだけを取りたい場合**  
  → `If tbls.Item(t).id = "targetTable"` のように `id` や `className` でフィルタする。  

- **表のヘッダだけを取りたい場合**  
  → `<th>` だけを抽出するロジックを残し、`<td>` をスキップする。  

- **HTML が日本語で文字化けする場合**  
  → サーバのエンコーディング（UTF-8 / Shift_JIS）を確認し、必要なら  
  `StrConv(http.ResponseBody, vbUnicode)` や `ADODB.Stream` を使って再エンコードする。  

---

✅ **まとめ**  
このスクリプトを既存の「cURL → VBA」スクリプトと組み合わせれば、  
ログイン後の HTML ページから自動的に表データを取り出し、  
Excel シートにそのまま書き込むことが可能です。  
WinHttp + MSHTML のみで完結し、外部インストール不要です。