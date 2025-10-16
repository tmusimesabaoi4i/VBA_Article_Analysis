# Chromium の DevTools を使った cURL → VBA 再現手順（.md）

## 手順（概観）

1. Chromium の DevTools（F12）を開く。  
2. Network タブで目的の操作（そのページにアクセスしてデータが表示される動作）を実行する。  
3. 目的の XHR / GET / POST リクエストを右クリック → **Copy → Copy as cURL** を選ぶ。  
4. 得られた **cURL コマンド**を下の VBA のヘルパー関数に渡す（あるいは自分でヘッダ / メソッド / Body を手で抜き出す）。  
5. **VBA（WinHttp）**で同じヘッダ・Cookie・ボディを設定して実行 → `ResponseText` を `MSHTML.HTMLDocument` で解析して必要データを取り出す。  

> 💡 **理由**  
> ブラウザ上で動くパスワードマネージャや SSO による認証はブラウザ側で処理されています。  
> DevTools のリクエストは「**認証済みセッションで投げられた正確なリクエスト**」なので、それをそのまま再現すればサーバは正常に応答します。

---

## すぐ使える VBA コード（改良版）

**やること：**  
- 「Copy as cURL」で取ったリクエストを VBA に写経  
- **ログイン POST** → **ターゲットページ GET**  
- 取得した HTML から **`<title>` を抽出して A1 に出力**

> 事前に、VBA の参照設定で **「Microsoft HTML Object Library」** を有効にしてください。  
> （VBE → ツール → 参照設定）

```vb
' ===============================================
' WinHttp + MSHTML サンプル（cURL再現 ＆ タイトルをA1に出力）
' ===============================================

' ▼使い方
' 1) DevTools で「Copy as cURL」を取得
' 2) 下の urlLogin / urlAfterLogin / headers() / body を自分の環境に合わせてコピペ
' 3) Run: Example_LoginAndGetTitle

Option Explicit

' 汎用送信関数
Private Function SendHttpRequest( _
    ByVal method As String, _
    ByVal url As String, _
    ByVal headers As Variant, _
    Optional ByVal body As String = "" _
) As Object
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open UCase$(method), url, False

    Dim i As Long
    If Not IsEmpty(headers) Then
        For i = LBound(headers) To UBound(headers)
            Dim parts() As String
            parts = Split(headers(i), ":", 2) ' "Name: Value"
            If UBound(parts) >= 1 Then
                http.setRequestHeader Trim$(parts(0)), Trim$(parts(1))
            End If
        Next i
    End If

    If Len(body) > 0 Then
        http.Send body
    Else
        http.Send
    End If

    If http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + 1000, "SendHttpRequest", _
                  "HTTP Error: " & http.Status & " " & http.StatusText
    End If

    Set SendHttpRequest = http ' レスポンスと Cookie を保持したまま返す
End Function

' タイトル抽出（MSHTMLを使用）
Private Function ExtractHtmlTitle(ByVal htmlText As String) As String
    Dim doc As New MSHTML.HTMLDocument
    ' HTMLDocument.Title を使うには Open/Write/Close で全文書き込みが安全
    doc.Open
    doc.Write htmlText
    doc.Close
    ExtractHtmlTitle = doc.Title
End Function

' 実行例：ログイン → ターゲットページ取得 → タイトルをA1へ
Public Sub Example_LoginAndGetTitle()
    On Error GoTo EH

    Dim urlLogin As String
    Dim urlAfterLogin As String
    Dim headersLogin As Variant
    Dim bodyLogin As String

    ' ====== ▼▼ ここを「Copy as cURL」から埋める ▼▼ ======
    ' 例）POST でログイン（フォームやAPIに合わせて編集）
    urlLogin = "https://internal.example.local/login"      ' cURLのURLに置換
    urlAfterLogin = "https://internal.example.local/home"  ' ログイン後に表示されるページ

    ' cURL の -H 'Name: value' をそのまま追加（必要なものだけでOK）
    headersLogin = Array( _
        "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8", _
        "Content-Type: application/x-www-form-urlencoded", _
        "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) ..." _
        ' "Cookie: xxx" は基本不要（ログイン前）。CSRF等で要求される場合のみ
    )

    ' cURL の --data-raw '...' / -d '...' をここへ
    bodyLogin = "username=YOUR_USER&password=YOUR_PASS" ' 例：URLエンコード済のフォーム

    ' ====== ▲▲ ここまで埋める ▲▲ ======

    ' 1) ログイン（POST）
    Dim http As Object
    Set http = SendHttpRequest("POST", urlLogin, headersLogin, bodyLogin)

    ' 2) 同じオブジェクトで遷移先（ホームやダッシュボード）を GET
    '    ※同じ http オブジェクトを使うと Cookie が保持される
    http.Open "GET", urlAfterLogin, False
    ' ログイン後の GET では通常ヘッダ最小限でOK（必要なら Accept などを再セット）
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) ..."
    http.Send

    If http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + 1001, "Example_LoginAndGetTitle", _
                  "After-login GET failed: " & http.Status & " " & http.StatusText
    End If

    Dim html As String
    html = http.ResponseText

    Dim titleText As String
    titleText = ExtractHtmlTitle(html)

    ' 結果を A1 に出力
    With ThisWorkbook.ActiveSheet
        .Range("A1").Value = titleText
        .Range("A1").EntireColumn.AutoFit
    End With

    MsgBox "完了：タイトルをA1に出力しました → " & titleText, vbInformation
    Exit Sub

EH:
    MsgBox "エラー: " & Err.Source & vbCrLf & Err.Description, vbExclamation
End Sub
```

> 🔧 **補足**  
> - ログインに CSRF トークンや追加ヘッダが必要な場合は、**最初に GET でログインページを取得 → HTML からトークン抽出 → POST** という 2段階にしてください。  
> - ログイン後に **API（JSON）** を叩いているサイトなら、`urlAfterLogin` でその API エンドポイントを GET/POST し、レスポンス JSON を解析するのが確実です。

---

## 「Copy as cURL」からヘッダを抜くコツ（初心者向け・超丁寧版）

### 1) まずは「どのリクエストをコピるか」を見極める
- Network タブには**大量の通信**が出ます。目的のデータが表示された直後に増えた **XHR / fetch / document** をクリック。  
- 右ペインの **Headers** を見て、`Request URL`・`Request Method`・`Status Code`・`Request Headers` をざっと確認。  
- これが「**実際にデータを取ってきたリクエスト**」なら、右クリック → **Copy → Copy as cURL**。

### 2) cURL のどの部分を VBA に写せばよいか
cURL 例（抜粋・例示）：
```
curl 'https://internal.example.local/api/getReport' \
  -X POST \
  -H 'Accept: application/json, text/javascript, */*; q=0.01' \
  -H 'Content-Type: application/json;charset=UTF-8' \
  -H 'Cookie: sessionid=abcdef123456; other=xxx' \
  --data-raw '{"param1":"value1","param2":123}'
```

VBA では以下に対応づけます：

- **URL** → `urlLogin` や `urlAfterLogin` にそのまま貼る  
- **メソッド**（`-X POST` / `-X GET`） → `SendHttpRequest("POST", ...)` の第1引数  
- **ヘッダ（-H 'Name: Value'）** → `headers = Array("Name: Value", ...)` の配列に**そのまま文字列で**入れる  
  - よく使う例：`Accept`, `Content-Type`, `Cookie`, `User-Agent`, `X-Requested-With`, `X-CSRF-Token` など  
  - **Cookie は超重要**：ログイン済みセッションを再現できます（※有効期限に注意）
- **ボディ（--data-raw '...' または -d '...'）** → 第4引数 `body` に**そのまま**貼る  
  - JSON の場合は `Content-Type: application/json` を忘れずに  
  - フォーム（例：`username=...&password=...`）なら `application/x-www-form-urlencoded`

> 🔎 **ポイント**  
> - まず **ブラウザと同じヘッダ**で試す → 動けば正解。  
> - 動かない場合、`Referer` や `Origin`、`X-CSRF-Token` を足すと通ることが多いです。  
> - 逆に、`sec-ch-ua` など**ブラウザ固有ヘッダ**は省略しても大抵動きます。最小限で通る組み合わせに整理しましょう。

### 3) 「Cookie をどうするか」
- ログイン直後の **API リクエスト**から cURL を取れば、`Cookie:` が含まれます。  
- その `Cookie:` をヘッダ配列にコピペするのが最短。  
- ただし Cookie は期限切れになります。**安定運用**したいなら：
  1. **VBA でログイン POST** を最初に実行（フォームや SSO 仕様に合わせる）  
  2. **同じ http オブジェクトで続けて GET/POST**（※これで Cookie を自動引き継ぎ）  
  3. 以後の処理は Cookie を手動で貼らなくてOK

### 4) 圧縮やエンコーディングで詰まったら
- `--compressed` が付いていても、WinHttp は **自動で解凍** してくれます（通常はそのままでOK）。  
- もし応答が文字化けする時は、レスポンスの `Content-Type`/`charset` を確認し、VBA 側で再エンコード（`ADODB.Stream` など）を検討。

### 5) 最小構成でまず動かす
- `Accept`, `Content-Type`, `User-Agent`, `Cookie`（必要なら）だけで**まず実行** → 通るか確認  
- 通らなければ、`Referer`, `Origin`, `X-CSRF-Token` を**一個ずつ**足して再テスト  
- それでもダメなら **リクエストURL／クエリ文字列の差異** や **サーバ側のエラーメッセージ** をチェック

---

## 解析・定期実行について

- **HTML の解析**：  
  `MSHTML.HTMLDocument` に全文を書き込み（`doc.Open → doc.Write → doc.Close`）、  
  `doc.Title` / `doc.getElementById(...)` / `doc.querySelector(...)` などで抽出できます。

- **JSON の解析**：  
  文字列操作（`InStr`, `Split`, 正規表現）でも可能ですが、扱うキーが多い場合は  
  軽量な JSON パーサ（VBA-JSON など）を導入すると楽です（導入が難しい環境ではまずは文字列処理でOK）。

- **定期実行**：  
  - Excel の `Application.OnTime` で **分・時単位の定期実行**  
  - あるいは Windows タスク スケジューラから、バッチで Excel を起動し対象ブックのマクロを呼ぶ  
  - 外部インストールなしで運用できます（社内ポリシー要確認）

---

## セキュリティの注意

- パスワードの**直書き禁止**。可能ならサービスアカウントやトークン方式に。  
- 社内の情報セキュリティポリシーに従う（自動取得・スクレイピングの可否、頻度制限）。  
- サーバ負荷に配慮（キャッシュ・間隔）。ログとエラー処理を実装して保守性を確保。

---

### ✅ まとめ

DevTools の「Copy as cURL」を出発点に、**WinHttp + MSHTML** を使えば、  
Chromium 上で認証済みの通信を **VBA だけで再現**し、データ抽出や自動化が可能です。  
上のサンプルをコピペして URL／ヘッダ／ボディを埋めれば、そのまま動きます。