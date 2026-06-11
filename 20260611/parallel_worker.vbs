Option Explicit

' ============================================================
' parallel_worker.vbs
'
' 親VBAから cscript.exe で並列起動されるワーカー。
'
' 引数:
'   0 jobPath
'   1 rootDir
'   2 proxyRaw
'   3 workerNo
' ============================================================

Const DIR_ZIP = "ZIP"
Const DIR_EXTRACT = "EXTRACT"
Const DIR_PDF = "PDF"
Const DIR_HTML = "HTML"

Const WD_EXPORT_FORMAT_PDF = 17
Const WD_FORMAT_FILTERED_HTML = 10

Dim gWorkerNo
Dim gFso

Set gFso = CreateObject("Scripting.FileSystemObject")

Main

Sub Main()

    Dim jobPath
    Dim rootDir
    Dim proxyRaw
    Dim workDir
    Dim resultPath
    Dim donePath

    Dim zipDir
    Dim extractDir
    Dim pdfDir
    Dim htmlDir

    Dim wordApp
    Dim jobText
    Dim lines
    Dim i
    Dim line
    Dim parts
    Dim rowNo
    Dim url

    Dim downloadedPath
    Dim extractPath
    Dim pdfResult
    Dim htmlResult
    Dim resultText
    Dim outputText

    If WScript.Arguments.Count < 4 Then
        WScript.Quit 1
    End If

    jobPath = WScript.Arguments(0)
    rootDir = NormalizeFolderPath(WScript.Arguments(1))
    proxyRaw = WScript.Arguments(2)
    gWorkerNo = CStr(WScript.Arguments(3))

    workDir = gFso.GetParentFolderName(jobPath)

    resultPath = CombinePath(workDir, "result_" & Right("000" & gWorkerNo, 3) & ".tsv")
    donePath = CombinePath(workDir, "done_" & Right("000" & gWorkerNo, 3) & ".txt")

    zipDir = CombinePath(rootDir, DIR_ZIP)
    extractDir = CombinePath(rootDir, DIR_EXTRACT)
    pdfDir = CombinePath(rootDir, DIR_PDF)
    htmlDir = CombinePath(rootDir, DIR_HTML)

    EnsureFolderExists zipDir
    EnsureFolderExists extractDir
    EnsureFolderExists pdfDir
    EnsureFolderExists htmlDir

    Set wordApp = Nothing

    On Error Resume Next
    Set wordApp = CreateObject("Word.Application")
    If Err.Number <> 0 Then
        Err.Clear
        Set wordApp = Nothing
    End If
    On Error GoTo 0

    If Not wordApp Is Nothing Then
        On Error Resume Next
        wordApp.Visible = False
        wordApp.DisplayAlerts = 0
        On Error GoTo 0
    End If

    jobText = ReadTextUtf8(jobPath)
    jobText = Replace(jobText, vbCrLf, vbLf)
    jobText = Replace(jobText, vbCr, vbLf)

    lines = Split(jobText, vbLf)

    outputText = ""

    For i = LBound(lines) To UBound(lines)

        line = Trim(lines(i))

        If Len(line) > 0 Then

            parts = Split(line, vbTab)

            If UBound(parts) >= 1 Then

                rowNo = Replace(parts(0), ChrW(&HFEFF), "")
                url = parts(1)

                downloadedPath = ""
                extractPath = ""
                pdfResult = ""
                htmlResult = ""
                resultText = ""

                resultText = DownloadOneFile(url, zipDir, proxyRaw, downloadedPath)

                If resultText = "OK" Then
                    resultText = ProcessDownloadedFile( _
                                    downloadedPath, _
                                    extractDir, _
                                    pdfDir, _
                                    htmlDir, _
                                    wordApp, _
                                    extractPath, _
                                    pdfResult, _
                                    htmlResult _
                                 )
                Else
                    pdfResult = ""
                    htmlResult = ""
                    extractPath = ""
                End If

                outputText = outputText & _
                    CleanField(rowNo) & vbTab & _
                    CleanField(url) & vbTab & _
                    CleanField(downloadedPath) & vbTab & _
                    CleanField(extractPath) & vbTab & _
                    CleanField(pdfResult) & vbTab & _
                    CleanField(htmlResult) & vbTab & _
                    CleanField(resultText) & vbCrLf

            End If

        End If

    Next

    On Error Resume Next
    If Not wordApp Is Nothing Then
        wordApp.Quit
    End If
    On Error GoTo 0

    WriteTextUtf8 resultPath, outputText
    WriteTextUtf8 donePath, "done"

End Sub

' ============================================================
' ダウンロード後処理
' ============================================================

Function ProcessDownloadedFile( _
        downloadedPath, _
        extractRootDir, _
        pdfRootDir, _
        htmlRootDir, _
        wordApp, _
        ByRef extractPath, _
        ByRef pdfResult, _
        ByRef htmlResult _
    )

    On Error Resume Next

    Dim baseName
    Dim thisExtractDir
    Dim thisPdfDir
    Dim pdfCount
    Dim htmlCount
    Dim onePdfPath
    Dim oneHtmlPath

    If Len(downloadedPath) = 0 Then
        ProcessDownloadedFile = "対象ファイルなし"
        Exit Function
    End If

    baseName = SanitizeFileName(gFso.GetBaseName(downloadedPath))

    If IsZipFile(downloadedPath) Then

        thisExtractDir = UniqueFolderPath(CombinePath(extractRootDir, baseName))
        thisPdfDir = UniqueFolderPath(CombinePath(pdfRootDir, baseName))

        EnsureFolderExists thisExtractDir
        EnsureFolderExists thisPdfDir
        EnsureFolderExists htmlRootDir

        ExtractZipToFolder downloadedPath, thisExtractDir

        extractPath = thisExtractDir

        pdfCount = ConvertFolderFilesToPdf(thisExtractDir, thisPdfDir, wordApp)
        htmlCount = ConvertFolderWordFilesToBeautifulHtml(thisExtractDir, htmlRootDir, wordApp)

        pdfResult = CStr(pdfCount) & "件 PDF作成"
        htmlResult = CStr(htmlCount) & "件 HTML作成"

        ProcessDownloadedFile = "ZIP展開OK / PDF作成OK / HTML作成OK"

    Else

        EnsureFolderExists pdfRootDir
        EnsureFolderExists htmlRootDir

        onePdfPath = ConvertSingleFileToPdf(downloadedPath, pdfRootDir, wordApp)

        If Len(onePdfPath) > 0 Then
            pdfResult = onePdfPath
        Else
            pdfResult = "PDF作成対象外"
        End If

        oneHtmlPath = ConvertSingleWordFileToBeautifulHtml(downloadedPath, htmlRootDir, wordApp)

        If Len(oneHtmlPath) > 0 Then
            htmlResult = oneHtmlPath
        Else
            htmlResult = "HTML作成対象外"
        End If

        ProcessDownloadedFile = "PDF/HTML作成処理完了"

    End If

    If Err.Number <> 0 Then
        ProcessDownloadedFile = "後処理エラー: " & Err.Description
        Err.Clear
    End If

End Function

' ============================================================
' ダウンロード
' ============================================================

Function DownloadOneFile( _
        url, _
        downloadDir, _
        proxyRaw, _
        ByRef savedPath _
    )

    On Error Resume Next

    Dim landingUrl
    Dim cookieHeader
    Dim req
    Dim statusCode
    Dim fileName
    Dim contentDisposition
    Dim contentType
    Dim body

    landingUrl = GetLandingUrl(url)
    cookieHeader = FetchCookieFromLanding(landingUrl, proxyRaw)

    Set req = CreateObject("WinHttp.WinHttpRequest.5.1")

    ApplyProxy req, proxyRaw

    req.SetTimeouts 30000, 30000, 30000, 180000

    Err.Clear
    req.SetAutoLogonPolicy 0
    Err.Clear

    req.Open "GET", url, False

    req.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    req.SetRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
    req.SetRequestHeader "Accept-Language", "en-US,en;q=0.9,ja;q=0.8"
    req.SetRequestHeader "Connection", "keep-alive"
    req.SetRequestHeader "Upgrade-Insecure-Requests", "1"

    If Len(cookieHeader) > 0 Then
        req.SetRequestHeader "Cookie", cookieHeader
    End If

    req.Send

    If Err.Number <> 0 Then
        DownloadOneFile = "ダウンロードエラー: " & Err.Description
        Err.Clear
        Exit Function
    End If

    statusCode = CLng(req.Status)

    If statusCode < 200 Or statusCode >= 300 Then
        DownloadOneFile = "失敗 HTTP " & CStr(statusCode) & " " & req.StatusText
        Exit Function
    End If

    contentDisposition = SafeGetResponseHeader(req, "Content-Disposition")
    contentType = SafeGetResponseHeader(req, "Content-Type")

    fileName = FileNameFromContentDisposition(contentDisposition)

    If Len(fileName) = 0 Then
        fileName = FileNameFromUrl(url)
    End If

    If Len(fileName) = 0 Then
        fileName = "download_" & CurrentTimestamp()
    End If

    fileName = SanitizeFileName(fileName)

    If InStrRev(fileName, ".") = 0 Then
        fileName = fileName & ExtensionFromContentType(contentType)
    End If

    savedPath = UniqueFilePath(CombinePath(downloadDir, fileName))

    body = req.ResponseBody

    SaveBinaryToFile body, savedPath

    If Err.Number <> 0 Then
        DownloadOneFile = "保存エラー: " & Err.Description
        Err.Clear
    Else
        DownloadOneFile = "OK"
    End If

End Function

Function FetchCookieFromLanding(landingUrl, proxyRaw)

    On Error Resume Next

    Dim req
    Dim headers

    Set req = CreateObject("WinHttp.WinHttpRequest.5.1")

    ApplyProxy req, proxyRaw

    req.SetTimeouts 30000, 30000, 30000, 60000

    Err.Clear
    req.SetAutoLogonPolicy 0
    Err.Clear

    req.Open "GET", landingUrl, False

    req.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    req.SetRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
    req.SetRequestHeader "Accept-Language", "en-US,en;q=0.9,ja;q=0.8"
    req.SetRequestHeader "Connection", "keep-alive"

    req.Send

    If Err.Number <> 0 Then
        Err.Clear
        FetchCookieFromLanding = ""
        Exit Function
    End If

    headers = req.GetAllResponseHeaders
    FetchCookieFromLanding = ParseSetCookieHeaders(headers)

End Function

Function ParseSetCookieHeaders(headers)

    Dim lines
    Dim i
    Dim line
    Dim v
    Dim semiPos
    Dim result

    headers = Replace(headers, vbCrLf, vbLf)
    lines = Split(headers, vbLf)

    result = ""

    For i = LBound(lines) To UBound(lines)

        line = Trim(lines(i))

        If LCase(Left(line, 11)) = "set-cookie:" Then

            v = Trim(Mid(line, 12))

            semiPos = InStr(1, v, ";")

            If semiPos > 0 Then
                v = Left(v, semiPos - 1)
            End If

            If Len(v) > 0 Then
                If Len(result) > 0 Then result = result & "; "
                result = result & v
            End If

        End If

    Next

    ParseSetCookieHeaders = result

End Function

Function SafeGetResponseHeader(req, headerName)

    On Error Resume Next

    SafeGetResponseHeader = req.GetResponseHeader(headerName)

    If Err.Number <> 0 Then
        Err.Clear
        SafeGetResponseHeader = ""
    End If

End Function

Sub ApplyProxy(req, proxyRaw)

    Dim proxyText

    proxyText = BuildWinHttpProxyString(proxyRaw)

    If Len(proxyText) = 0 Then
        req.SetProxy 0
    Else
        req.SetProxy 2, proxyText
    End If

End Sub

Function BuildWinHttpProxyString(proxyRaw)

    Dim p

    p = Trim(proxyRaw)

    If Len(p) = 0 Then
        BuildWinHttpProxyString = ""
        Exit Function
    End If

    If InStr(p, "=") > 0 Or InStr(p, ";") > 0 Then
        BuildWinHttpProxyString = p
        Exit Function
    End If

    p = Replace(p, "http://", "", 1, -1, 1)
    p = Replace(p, "https://", "", 1, -1, 1)

    Do While Right(p, 1) = "/"
        p = Left(p, Len(p) - 1)
    Loop

    BuildWinHttpProxyString = "http=" & p & ";https=" & p

End Function

' ============================================================
' ZIP展開
' ============================================================

Sub ExtractZipToFolder(zipPath, destFolder)

    On Error Resume Next

    Dim shellApp
    Dim zipNs
    Dim destNs

    EnsureFolderExists destFolder

    Set shellApp = CreateObject("Shell.Application")
    Set zipNs = shellApp.NameSpace(zipPath)
    Set destNs = shellApp.NameSpace(destFolder)

    If zipNs Is Nothing Then Exit Sub
    If destNs Is Nothing Then Exit Sub

    destNs.CopyHere zipNs.Items, 4 + 16 + 512 + 1024

    WaitForExtraction destFolder

End Sub

Sub WaitForExtraction(destFolder)

    Dim i
    Dim c
    Dim lastC
    Dim stableCount

    lastC = -1
    stableCount = 0

    For i = 1 To 90

        WScript.Sleep 1000

        c = CountFilesRecursive(destFolder)

        If c > 0 And c = lastC Then
            stableCount = stableCount + 1
        Else
            stableCount = 0
        End If

        If stableCount >= 2 Then Exit For

        lastC = c

    Next

End Sub

Function CountFilesRecursive(folderPath)

    On Error Resume Next

    Dim folder
    Dim file
    Dim subFolder
    Dim count

    count = 0

    If Not gFso.FolderExists(folderPath) Then
        CountFilesRecursive = 0
        Exit Function
    End If

    Set folder = gFso.GetFolder(folderPath)

    count = folder.Files.Count

    For Each subFolder In folder.SubFolders
        count = count + CountFilesRecursive(CStr(subFolder.Path))
    Next

    CountFilesRecursive = count

End Function

Function IsZipFile(filePath)

    Dim ext

    ext = LCase(gFso.GetExtensionName(filePath))

    If ext = "zip" Then
        IsZipFile = True
        Exit Function
    End If

    IsZipFile = HasZipSignature(filePath)

End Function

Function HasZipSignature(filePath)

    On Error Resume Next

    Dim st
    Dim b

    HasZipSignature = False

    Set st = CreateObject("ADODB.Stream")

    st.Type = 1
    st.Open
    st.LoadFromFile filePath

    If st.Size >= 2 Then
        b = st.Read(2)
        If IsArray(b) Then
            If b(0) = 80 And b(1) = 75 Then
                HasZipSignature = True
            End If
        End If
    End If

    st.Close

    If Err.Number <> 0 Then
        Err.Clear
        HasZipSignature = False
    End If

End Function

' ============================================================
' PDF作成
' ============================================================

Function ConvertFolderFilesToPdf(sourceFolder, pdfFolder, wordApp)

    On Error Resume Next

    Dim folder
    Dim file
    Dim subFolder
    Dim count
    Dim pdfPath

    count = 0

    If Not gFso.FolderExists(sourceFolder) Then
        ConvertFolderFilesToPdf = 0
        Exit Function
    End If

    EnsureFolderExists pdfFolder

    Set folder = gFso.GetFolder(sourceFolder)

    For Each file In folder.Files
        pdfPath = ConvertSingleFileToPdf(CStr(file.Path), pdfFolder, wordApp)
        If Len(pdfPath) > 0 Then count = count + 1
    Next

    For Each subFolder In folder.SubFolders
        count = count + ConvertFolderFilesToPdf(CStr(subFolder.Path), pdfFolder, wordApp)
    Next

    ConvertFolderFilesToPdf = count

End Function

Function ConvertSingleFileToPdf(filePath, pdfFolder, wordApp)

    On Error Resume Next

    Dim ext
    Dim pdfPath

    If Not gFso.FileExists(filePath) Then
        ConvertSingleFileToPdf = ""
        Exit Function
    End If

    EnsureFolderExists pdfFolder

    ext = LCase(gFso.GetExtensionName(filePath))

    pdfPath = UniqueFilePath(CombinePath(pdfFolder, SanitizeFileName(gFso.GetBaseName(filePath)) & ".pdf"))

    Select Case ext

        Case "pdf"
            gFso.CopyFile filePath, pdfPath, True
            ConvertSingleFileToPdf = pdfPath

        Case "html", "htm", "txt", "rtf", "doc", "docx", "docm"
            If wordApp Is Nothing Then
                ConvertSingleFileToPdf = ""
            Else
                ConvertWordOpenableFileToPdf filePath, pdfPath, wordApp
                If gFso.FileExists(pdfPath) Then
                    ConvertSingleFileToPdf = pdfPath
                Else
                    ConvertSingleFileToPdf = ""
                End If
            End If

        Case "jpg", "jpeg", "png", "bmp", "gif", "tif", "tiff"
            If wordApp Is Nothing Then
                ConvertSingleFileToPdf = ""
            Else
                ConvertImageToPdf filePath, pdfPath, wordApp
                If gFso.FileExists(pdfPath) Then
                    ConvertSingleFileToPdf = pdfPath
                Else
                    ConvertSingleFileToPdf = ""
                End If
            End If

        Case Else
            ConvertSingleFileToPdf = ""

    End Select

    If Err.Number <> 0 Then
        Err.Clear
        ConvertSingleFileToPdf = ""
    End If

End Function

Sub ConvertWordOpenableFileToPdf(filePath, pdfPath, wordApp)

    On Error Resume Next

    Dim doc

    Set doc = wordApp.Documents.Open(filePath, False, True, False)

    If Not doc Is Nothing Then
        doc.ExportAsFixedFormat pdfPath, WD_EXPORT_FORMAT_PDF
        doc.Close False
    End If

    If Err.Number <> 0 Then Err.Clear

End Sub

Sub ConvertImageToPdf(imagePath, pdfPath, wordApp)

    On Error Resume Next

    Dim doc
    Dim img
    Dim availableWidth
    Dim availableHeight

    Set doc = wordApp.Documents.Add

    If doc Is Nothing Then Exit Sub

    doc.PageSetup.TopMargin = 36
    doc.PageSetup.BottomMargin = 36
    doc.PageSetup.LeftMargin = 36
    doc.PageSetup.RightMargin = 36

    availableWidth = doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin
    availableHeight = doc.PageSetup.PageHeight - doc.PageSetup.TopMargin - doc.PageSetup.BottomMargin

    Set img = doc.InlineShapes.AddPicture(imagePath, False, True)

    If Not img Is Nothing Then
        img.LockAspectRatio = True

        If img.Width > availableWidth Then
            img.Width = availableWidth
        End If

        If img.Height > availableHeight Then
            img.Height = availableHeight
        End If
    End If

    doc.ExportAsFixedFormat pdfPath, WD_EXPORT_FORMAT_PDF
    doc.Close False

    If Err.Number <> 0 Then Err.Clear

End Sub

' ============================================================
' HTML作成
' ============================================================

Function ConvertFolderWordFilesToBeautifulHtml(sourceFolder, htmlFolder, wordApp)

    On Error Resume Next

    Dim folder
    Dim file
    Dim subFolder
    Dim count
    Dim htmlPath

    count = 0

    If wordApp Is Nothing Then
        ConvertFolderWordFilesToBeautifulHtml = 0
        Exit Function
    End If

    If Not gFso.FolderExists(sourceFolder) Then
        ConvertFolderWordFilesToBeautifulHtml = 0
        Exit Function
    End If

    EnsureFolderExists htmlFolder

    Set folder = gFso.GetFolder(sourceFolder)

    For Each file In folder.Files

        htmlPath = ConvertSingleWordFileToBeautifulHtml(CStr(file.Path), htmlFolder, wordApp)

        If Len(htmlPath) > 0 Then
            count = count + 1
        End If

    Next

    For Each subFolder In folder.SubFolders
        count = count + ConvertFolderWordFilesToBeautifulHtml(CStr(subFolder.Path), htmlFolder, wordApp)
    Next

    ConvertFolderWordFilesToBeautifulHtml = count

End Function

Function ConvertSingleWordFileToBeautifulHtml(filePath, htmlFolder, wordApp)

    On Error Resume Next

    Dim ext
    Dim baseName
    Dim outHtmlPath
    Dim rawHtml
    Dim bodyHtml
    Dim prettyHtml

    If wordApp Is Nothing Then
        ConvertSingleWordFileToBeautifulHtml = ""
        Exit Function
    End If

    If Not gFso.FileExists(filePath) Then
        ConvertSingleWordFileToBeautifulHtml = ""
        Exit Function
    End If

    ext = LCase(gFso.GetExtensionName(filePath))

    If Not IsWordHtmlConvertibleExt(ext) Then
        ConvertSingleWordFileToBeautifulHtml = ""
        Exit Function
    End If

    EnsureFolderExists htmlFolder

    baseName = SanitizeFileName(gFso.GetBaseName(filePath))
    outHtmlPath = UniqueFilePath(CombinePath(htmlFolder, baseName & ".html"))

    Select Case ext

        Case "doc", "docx", "docm", "rtf", "txt"
            SaveWordOpenableFileAsFilteredHtml filePath, outHtmlPath, wordApp
            rawHtml = ReadTextUtf8(outHtmlPath)

        Case "html", "htm"
            rawHtml = ReadTextUtf8(filePath)

        Case Else
            ConvertSingleWordFileToBeautifulHtml = ""
            Exit Function

    End Select

    bodyHtml = ExtractBodyInnerHtml(rawHtml)

    If Len(Trim(bodyHtml)) = 0 Then
        bodyHtml = "<pre>" & HtmlEscape(ReadTextDefault(filePath)) & "</pre>"
    End If

    prettyHtml = BuildBeautifulHtmlDocument(baseName, gFso.GetFileName(filePath), bodyHtml)

    WriteTextUtf8 outHtmlPath, prettyHtml

    If gFso.FileExists(outHtmlPath) Then
        ConvertSingleWordFileToBeautifulHtml = outHtmlPath
    Else
        ConvertSingleWordFileToBeautifulHtml = ""
    End If

    If Err.Number <> 0 Then
        Err.Clear
        ConvertSingleWordFileToBeautifulHtml = ""
    End If

End Function

Function IsWordHtmlConvertibleExt(ext)

    Select Case LCase(ext)

        Case "doc", "docx", "docm", "rtf", "txt", "html", "htm"
            IsWordHtmlConvertibleExt = True

        Case Else
            IsWordHtmlConvertibleExt = False

    End Select

End Function

Sub SaveWordOpenableFileAsFilteredHtml(filePath, outHtmlPath, wordApp)

    On Error Resume Next

    Dim doc

    Set doc = wordApp.Documents.Open(filePath, False, True, False)

    If Not doc Is Nothing Then
        doc.WebOptions.Encoding = 65001
        doc.SaveAs2 outHtmlPath, WD_FORMAT_FILTERED_HTML
        doc.Close False
    End If

    If Err.Number <> 0 Then Err.Clear

End Sub

Function BuildBeautifulHtmlDocument(titleText, sourceFileName, bodyHtml)

    Dim s

    s = ""
    s = s & "<!doctype html>" & vbCrLf
    s = s & "<html lang=""ja"">" & vbCrLf
    s = s & "<head>" & vbCrLf
    s = s & "<meta charset=""utf-8"">" & vbCrLf
    s = s & "<meta name=""viewport"" content=""width=device-width, initial-scale=1"">" & vbCrLf
    s = s & "<title>" & HtmlEscape(titleText) & "</title>" & vbCrLf
    s = s & BeautifulHtmlCss() & vbCrLf
    s = s & "</head>" & vbCrLf
    s = s & "<body>" & vbCrLf
    s = s & "<main class=""page"">" & vbCrLf
    s = s & "<header class=""doc-header"">" & vbCrLf
    s = s & "<div class=""label"">WORD HTML VIEW</div>" & vbCrLf
    s = s & "<h1>" & HtmlEscape(titleText) & "</h1>" & vbCrLf
    s = s & "<div class=""meta"">Source: " & HtmlEscape(sourceFileName) & "</div>" & vbCrLf
    s = s & "</header>" & vbCrLf
    s = s & "<article class=""doc-body"">" & vbCrLf
    s = s & bodyHtml & vbCrLf
    s = s & "</article>" & vbCrLf
    s = s & "</main>" & vbCrLf
    s = s & "</body>" & vbCrLf
    s = s & "</html>" & vbCrLf

    BuildBeautifulHtmlDocument = s

End Function

Function BeautifulHtmlCss()

    Dim s

    s = ""
    s = s & "<style>" & vbCrLf
    s = s & "html{background:#f3f4f6;color:#111827;}" & vbCrLf
    s = s & "body{margin:0;font-family:-apple-system,BlinkMacSystemFont,'Yu Gothic','Meiryo','Segoe UI',sans-serif;line-height:1.85;}" & vbCrLf
    s = s & ".page{max-width:980px;margin:32px auto;padding:40px;background:#fff;border-radius:18px;box-shadow:0 12px 36px rgba(15,23,42,.12);}" & vbCrLf
    s = s & ".doc-header{border-bottom:1px solid #e5e7eb;margin-bottom:28px;padding-bottom:18px;}" & vbCrLf
    s = s & ".label{font-size:12px;letter-spacing:.14em;color:#6b7280;font-weight:700;margin-bottom:8px;}" & vbCrLf
    s = s & "h1{font-size:26px;line-height:1.35;margin:0 0 8px 0;color:#0f172a;}" & vbCrLf
    s = s & ".meta{font-size:13px;color:#6b7280;}" & vbCrLf
    s = s & ".doc-body{font-size:15.5px;}" & vbCrLf
    s = s & ".doc-body p{margin:.65em 0;}" & vbCrLf
    s = s & ".doc-body h1,.doc-body h2,.doc-body h3{line-height:1.45;margin-top:1.4em;color:#111827;}" & vbCrLf
    s = s & ".doc-body table{border-collapse:collapse;width:100%;margin:18px 0;font-size:14px;}" & vbCrLf
    s = s & ".doc-body th,.doc-body td{border:1px solid #d1d5db;padding:8px 10px;vertical-align:top;}" & vbCrLf
    s = s & ".doc-body th{background:#f9fafb;font-weight:700;}" & vbCrLf
    s = s & ".doc-body img{max-width:100%;height:auto;display:block;margin:16px auto;}" & vbCrLf
    s = s & ".doc-body pre{white-space:pre-wrap;background:#f9fafb;border:1px solid #e5e7eb;padding:16px;border-radius:12px;}" & vbCrLf
    s = s & ".doc-body a{color:#2563eb;text-decoration:none;}" & vbCrLf
    s = s & ".doc-body a:hover{text-decoration:underline;}" & vbCrLf
    s = s & "@media print{html{background:#fff}.page{margin:0;box-shadow:none;border-radius:0}}" & vbCrLf
    s = s & "</style>" & vbCrLf

    BeautifulHtmlCss = s

End Function

Function ExtractBodyInnerHtml(html)

    Dim lowerHtml
    Dim bodyStart
    Dim bodyStartClose
    Dim bodyEnd

    lowerHtml = LCase(html)

    bodyStart = InStr(1, lowerHtml, "<body")

    If bodyStart = 0 Then
        ExtractBodyInnerHtml = html
        Exit Function
    End If

    bodyStartClose = InStr(bodyStart, lowerHtml, ">")

    If bodyStartClose = 0 Then
        ExtractBodyInnerHtml = html
        Exit Function
    End If

    bodyEnd = InStr(bodyStartClose + 1, lowerHtml, "</body>")

    If bodyEnd = 0 Then
        ExtractBodyInnerHtml = Mid(html, bodyStartClose + 1)
    Else
        ExtractBodyInnerHtml = Mid(html, bodyStartClose + 1, bodyEnd - bodyStartClose - 1)
    End If

End Function

' ============================================================
' URL / ファイル名
' ============================================================

Function GetLandingUrl(url)

    Dim schemeEnd
    Dim rest
    Dim slashPos
    Dim scheme
    Dim host

    schemeEnd = InStr(1, url, "://")

    If schemeEnd = 0 Then
        GetLandingUrl = url
        Exit Function
    End If

    scheme = Left(url, schemeEnd + 2)
    rest = Mid(url, schemeEnd + 3)

    slashPos = InStr(1, rest, "/")

    If slashPos > 0 Then
        host = Left(rest, slashPos - 1)
    Else
        host = rest
    End If

    GetLandingUrl = scheme & host & "/"

End Function

Function FileNameFromUrl(url)

    Dim u
    Dim qPos
    Dim hashPos
    Dim slashPos

    u = url

    hashPos = InStr(1, u, "#")
    If hashPos > 0 Then u = Left(u, hashPos - 1)

    qPos = InStr(1, u, "?")
    If qPos > 0 Then u = Left(u, qPos - 1)

    slashPos = InStrRev(u, "/")

    If slashPos > 0 Then
        FileNameFromUrl = Mid(u, slashPos + 1)
    Else
        FileNameFromUrl = u
    End If

    FileNameFromUrl = UrlDecodeLight(FileNameFromUrl)

End Function

Function FileNameFromContentDisposition(cd)

    Dim p
    Dim s
    Dim semiPos

    If Len(cd) = 0 Then
        FileNameFromContentDisposition = ""
        Exit Function
    End If

    p = InStr(1, cd, "filename*=", 1)

    If p > 0 Then

        s = Mid(cd, p + Len("filename*="))

        semiPos = InStr(1, s, ";")
        If semiPos > 0 Then s = Left(s, semiPos - 1)

        s = Replace(s, "UTF-8''", "", 1, -1, 1)
        s = Replace(s, """", "")

        FileNameFromContentDisposition = UrlDecodeLight(Trim(s))
        Exit Function

    End If

    p = InStr(1, cd, "filename=", 1)

    If p > 0 Then

        s = Mid(cd, p + Len("filename="))

        semiPos = InStr(1, s, ";")
        If semiPos > 0 Then s = Left(s, semiPos - 1)

        s = Replace(s, """", "")

        FileNameFromContentDisposition = Trim(s)
        Exit Function

    End If

    FileNameFromContentDisposition = ""

End Function

Function ExtensionFromContentType(contentType)

    Dim ct

    ct = LCase(contentType)

    If InStr(ct, "zip") > 0 Then
        ExtensionFromContentType = ".zip"
    ElseIf InStr(ct, "pdf") > 0 Then
        ExtensionFromContentType = ".pdf"
    ElseIf InStr(ct, "html") > 0 Then
        ExtensionFromContentType = ".html"
    ElseIf InStr(ct, "xml") > 0 Then
        ExtensionFromContentType = ".xml"
    ElseIf InStr(ct, "json") > 0 Then
        ExtensionFromContentType = ".json"
    ElseIf InStr(ct, "plain") > 0 Then
        ExtensionFromContentType = ".txt"
    ElseIf InStr(ct, "png") > 0 Then
        ExtensionFromContentType = ".png"
    ElseIf InStr(ct, "jpeg") > 0 Or InStr(ct, "jpg") > 0 Then
        ExtensionFromContentType = ".jpg"
    Else
        ExtensionFromContentType = ".bin"
    End If

End Function

Function SanitizeFileName(name)

    Dim badChars
    Dim i

    badChars = Array("<", ">", ":", """", "/", "\", "|", "?", "*")

    For i = LBound(badChars) To UBound(badChars)
        name = Replace(name, badChars(i), "_")
    Next

    name = Trim(name)

    Do While Len(name) > 0 And Right(name, 1) = "."
        name = Left(name, Len(name) - 1)
    Loop

    If Len(name) = 0 Then name = "page"

    SanitizeFileName = name

End Function

Function UrlDecodeLight(s)

    s = Replace(s, "%20", " ")
    s = Replace(s, "%2D", "-")
    s = Replace(s, "%2d", "-")
    s = Replace(s, "%5F", "_")
    s = Replace(s, "%5f", "_")
    s = Replace(s, "%2E", ".")
    s = Replace(s, "%2e", ".")

    UrlDecodeLight = s

End Function

' ============================================================
' ファイル・フォルダ共通
' ============================================================

Sub SaveBinaryToFile(body, filePath)

    On Error Resume Next

    Dim stream

    Set stream = CreateObject("ADODB.Stream")

    stream.Type = 1
    stream.Open
    stream.Write body
    stream.SaveToFile filePath, 2
    stream.Close

    If Err.Number <> 0 Then Err.Clear

End Sub

Function NormalizeFolderPath(folderPath)

    folderPath = Trim(folderPath)
    folderPath = Replace(folderPath, """", "")

    Do While Len(folderPath) > 3 And _
        (Right(folderPath, 1) = "\" Or Right(folderPath, 1) = "/")
        folderPath = Left(folderPath, Len(folderPath) - 1)
    Loop

    NormalizeFolderPath = folderPath

End Function

Function CombinePath(folderPath, fileName)

    If Right(folderPath, 1) = "\" Or Right(folderPath, 1) = "/" Then
        CombinePath = folderPath & fileName
    Else
        CombinePath = folderPath & "\" & fileName
    End If

End Function

Sub EnsureFolderExists(folderPath)

    On Error Resume Next

    Dim parentPath

    folderPath = NormalizeFolderPath(folderPath)

    If Len(folderPath) = 0 Then Exit Sub
    If gFso.FolderExists(folderPath) Then Exit Sub

    parentPath = gFso.GetParentFolderName(folderPath)

    If Len(parentPath) > 0 Then
        If Not gFso.FolderExists(parentPath) Then
            EnsureFolderExists parentPath
        End If
    End If

    If Not gFso.FolderExists(folderPath) Then
        gFso.CreateFolder folderPath
    End If

    If Err.Number <> 0 Then Err.Clear

End Sub

Function UniqueFilePath(filePath)

    Dim folderPath
    Dim baseName
    Dim ext
    Dim i
    Dim candidate

    If Not gFso.FileExists(filePath) Then
        UniqueFilePath = filePath
        Exit Function
    End If

    folderPath = gFso.GetParentFolderName(filePath)
    baseName = gFso.GetBaseName(filePath)
    ext = gFso.GetExtensionName(filePath)

    For i = 1 To 9999

        If Len(ext) > 0 Then
            candidate = folderPath & "\" & baseName & "_" & Right("000" & CStr(i), 3) & "." & ext
        Else
            candidate = folderPath & "\" & baseName & "_" & Right("000" & CStr(i), 3)
        End If

        If Not gFso.FileExists(candidate) Then
            UniqueFilePath = candidate
            Exit Function
        End If

    Next

    UniqueFilePath = filePath

End Function

Function UniqueFolderPath(folderPath)

    Dim i
    Dim candidate

    If Not gFso.FolderExists(folderPath) Then
        UniqueFolderPath = folderPath
        Exit Function
    End If

    For i = 1 To 9999

        candidate = folderPath & "_" & Right("000" & CStr(i), 3)

        If Not gFso.FolderExists(candidate) Then
            UniqueFolderPath = candidate
            Exit Function
        End If

    Next

    UniqueFolderPath = folderPath

End Function

Function ReadTextUtf8(filePath)

    On Error Resume Next

    Dim stream

    Set stream = CreateObject("ADODB.Stream")

    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadTextUtf8 = stream.ReadText
    stream.Close

    If Err.Number <> 0 Then
        Err.Clear
        ReadTextUtf8 = ReadTextDefault(filePath)
    End If

End Function

Function ReadTextDefault(filePath)

    On Error Resume Next

    Dim ts

    Set ts = gFso.OpenTextFile(filePath, 1, False, -2)

    ReadTextDefault = ts.ReadAll
    ts.Close

    If Err.Number <> 0 Then
        Err.Clear
        ReadTextDefault = ""
    End If

End Function

Sub WriteTextUtf8(filePath, text)

    On Error Resume Next

    Dim stream

    Set stream = CreateObject("ADODB.Stream")

    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.SaveToFile filePath, 2
    stream.Close

    If Err.Number <> 0 Then Err.Clear

End Sub

Function HtmlEscape(s)

    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&#39;")

    HtmlEscape = s

End Function

Function CleanField(s)

    s = CStr(s)
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")

    CleanField = s

End Function

Function CurrentTimestamp()

    Dim d

    d = Now

    CurrentTimestamp = _
        Year(d) & _
        Right("0" & Month(d), 2) & _
        Right("0" & Day(d), 2) & "_" & _
        Right("0" & Hour(d), 2) & _
        Right("0" & Minute(d), 2) & _
        Right("0" & Second(d), 2)

End Function