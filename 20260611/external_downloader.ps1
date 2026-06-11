param(
    [string]$Mode = "Master",
    [string]$JobPath = "",
    [string]$RootDir = "",
    [string]$ProxyRaw = "",
    [int]$MaxWorkers = 3,
    [int]$WorkerNo = 1,
    [string]$WorkDir = ""
)

$ErrorActionPreference = "Continue"

$DIR_ZIP = "ZIP"
$DIR_EXTRACT = "EXTRACT"
$DIR_PDF = "PDF"
$DIR_HTML = "HTML"
$DIR_HTML_COMBINE = "HTML_COMBINE"
$DIR_WORK = "WORK"

$script:CookieCache = @{}
$script:WordApp = $null
$script:ProgressPath = $null
$script:WorkerLabel = "{0:000}" -f $WorkerNo

# ============================================================
# Entry
# ============================================================

if ($Mode -eq "Master") {
    Invoke-Master
}
elseif ($Mode -eq "Worker") {
    Invoke-Worker
}
else {
    Write-Host "[ERROR] Unknown Mode: $Mode"
    exit 1
}

# ============================================================
# Master
# ============================================================

function Invoke-Master {

    if ([string]::IsNullOrWhiteSpace($JobPath)) {
        Write-Host "[ERROR] JobPath is empty."
        exit 1
    }

    if ([string]::IsNullOrWhiteSpace($RootDir)) {
        Write-Host "[ERROR] RootDir is empty."
        exit 1
    }

    $RootDir = Normalize-FolderPath $RootDir

    Ensure-Dir $RootDir
    Ensure-Dir (Join-Path $RootDir $DIR_ZIP)
    Ensure-Dir (Join-Path $RootDir $DIR_EXTRACT)
    Ensure-Dir (Join-Path $RootDir $DIR_PDF)
    Ensure-Dir (Join-Path $RootDir $DIR_HTML)
    Ensure-Dir (Join-Path $RootDir $DIR_HTML_COMBINE)
    Ensure-Dir (Join-Path $RootDir $DIR_WORK)

    $runDir = Join-Path (Join-Path $RootDir $DIR_WORK) ("run_" + (Get-Date -Format "yyyyMMdd_HHmmss"))
    Ensure-Dir $runDir

    Write-Host "============================================================"
    Write-Host "[MASTER] START"
    Write-Host "[MASTER] JobPath    = $JobPath"
    Write-Host "[MASTER] RootDir    = $RootDir"
    Write-Host "[MASTER] RunDir     = $runDir"
    Write-Host "[MASTER] ProxyRaw   = $ProxyRaw"
    Write-Host "[MASTER] MaxWorkers = $MaxWorkers"
    Write-Host "============================================================"

    $jobs = @(Get-Content -LiteralPath $JobPath -Encoding UTF8 | Where-Object { $_.Trim().Length -gt 0 })

    if ($jobs.Count -eq 0) {
        Write-Host "[MASTER][ERROR] No jobs."
        exit 1
    }

    if ($MaxWorkers -lt 1) {
        $MaxWorkers = 1
    }

    if ($MaxWorkers -gt 8) {
        $MaxWorkers = 8
    }

    if ($MaxWorkers -gt $jobs.Count) {
        $MaxWorkers = $jobs.Count
    }

    Write-Host "[MASTER] JobCount    = $($jobs.Count)"
    Write-Host "[MASTER] WorkerCount = $MaxWorkers"

    $workerJobPaths = @()

    for ($w = 1; $w -le $MaxWorkers; $w++) {
        $workerJobPaths += (Join-Path $runDir ("job_{0:000}.tsv" -f $w))
    }

    $buffers = @{}
    for ($w = 1; $w -le $MaxWorkers; $w++) {
        $buffers[$w] = New-Object System.Collections.Generic.List[string]
    }

    for ($i = 0; $i -lt $jobs.Count; $i++) {
        $w = ($i % $MaxWorkers) + 1
        $buffers[$w].Add($jobs[$i])
    }

    for ($w = 1; $w -le $MaxWorkers; $w++) {
        [System.IO.File]::WriteAllLines($workerJobPaths[$w - 1], $buffers[$w], [System.Text.Encoding]::UTF8)
    }

    $processes = @{}

    for ($w = 1; $w -le $MaxWorkers; $w++) {

        $workerJobPath = $workerJobPaths[$w - 1]

        if (-not (Test-Path -LiteralPath $workerJobPath)) {
            continue
        }

        $arg = ""
        $arg += "-NoProfile -ExecutionPolicy Bypass "
        $arg += "-File `"$PSCommandPath`" "
        $arg += "-Mode Worker "
        $arg += "-JobPath `"$workerJobPath`" "
        $arg += "-RootDir `"$RootDir`" "
        $arg += "-ProxyRaw `"$ProxyRaw`" "
        $arg += "-WorkerNo $w "
        $arg += "-WorkDir `"$runDir`" "

        Write-Host ("[MASTER] START WORKER {0:000}" -f $w)

        $p = Start-Process `
            -FilePath "powershell.exe" `
            -ArgumentList $arg `
            -PassThru `
            -WindowStyle Hidden

        $processes[$w] = $p

        Start-Sleep -Milliseconds 300
    }

    Monitor-Workers -RunDir $runDir -WorkerCount $MaxWorkers -Processes $processes

    Write-Host "[MASTER] Import/Combine phase started."

    $htmlDir = Join-Path $RootDir $DIR_HTML
    $combineDir = Join-Path $RootDir $DIR_HTML_COMBINE

    $combinedCount = Combine-HtmlFilesByFive -HtmlDir $htmlDir -OutputDir $combineDir

    Write-Host "[MASTER] Combined HTML count: $combinedCount"
    Write-Host "[MASTER] RunDir: $runDir"
    Write-Host "[MASTER] DONE"
}

function Monitor-Workers {
    param(
        [string]$RunDir,
        [int]$WorkerCount,
        [hashtable]$Processes
    )

    $lastProgress = @{}
    $lastSummary = Get-Date

    while ($true) {

        $doneCount = @(Get-ChildItem -LiteralPath $RunDir -Filter "done_*.txt" -ErrorAction SilentlyContinue).Count
        $runningCount = 0

        foreach ($key in $Processes.Keys) {
            $p = $Processes[$key]
            if (-not $p.HasExited) {
                $runningCount++
            }
        }

        for ($w = 1; $w -le $WorkerCount; $w++) {

            $label = "{0:000}" -f $w
            $progressFile = Join-Path $RunDir "progress_$label.txt"

            if (Test-Path -LiteralPath $progressFile) {

                $text = ""
                try {
                    $text = (Get-Content -LiteralPath $progressFile -Encoding UTF8 -Raw).Trim()
                }
                catch {
                    $text = ""
                }

                if ($text.Length -gt 0) {
                    if (-not $lastProgress.ContainsKey($label) -or $lastProgress[$label] -ne $text) {
                        Write-Host "[WORKER $label] $text"
                        $lastProgress[$label] = $text
                    }
                }
            }
        }

        if (((Get-Date) - $lastSummary).TotalSeconds -ge 5) {
            Write-Host "[MASTER] summary done=$doneCount/$WorkerCount running=$runningCount"
            $lastSummary = Get-Date
        }

        if ($doneCount -ge $WorkerCount) {
            break
        }

        if ($runningCount -eq 0) {
            Write-Host "[MASTER][WARN] All worker processes exited, but done files are insufficient. done=$doneCount/$WorkerCount"
            break
        }

        Start-Sleep -Milliseconds 700
    }
}

# ============================================================
# Worker
# ============================================================

function Invoke-Worker {

    $script:WorkerLabel = "{0:000}" -f $WorkerNo

    if ([string]::IsNullOrWhiteSpace($WorkDir)) {
        $WorkDir = Split-Path -Parent $JobPath
    }

    $script:ProgressPath = Join-Path $WorkDir "progress_$script:WorkerLabel.txt"

    $RootDir = Normalize-FolderPath $RootDir

    $zipDir = Join-Path $RootDir $DIR_ZIP
    $extractDir = Join-Path $RootDir $DIR_EXTRACT
    $pdfDir = Join-Path $RootDir $DIR_PDF
    $htmlDir = Join-Path $RootDir $DIR_HTML

    Ensure-Dir $zipDir
    Ensure-Dir $extractDir
    Ensure-Dir $pdfDir
    Ensure-Dir $htmlDir

    $resultPath = Join-Path $WorkDir "result_$script:WorkerLabel.tsv"
    $donePath = Join-Path $WorkDir "done_$script:WorkerLabel.txt"

    Write-ProgressFile -Stage "START" -Row "" -Message "job=$JobPath"

    $lines = @(Get-Content -LiteralPath $JobPath -Encoding UTF8 | Where-Object { $_.Trim().Length -gt 0 })
    $resultLines = New-Object System.Collections.Generic.List[string]

    foreach ($line in $lines) {

        $parts = $line -split "`t", 2

        if ($parts.Count -lt 2) {
            continue
        }

        $rowNo = ($parts[0] -replace "^\uFEFF", "").Trim()
        $url = $parts[1].Trim()

        $downloadedPath = ""
        $extractPath = ""
        $pdfResult = ""
        $htmlResult = ""
        $resultText = ""

        Write-ProgressFile -Stage "DOWNLOAD_START" -Row $rowNo -Message $url

        try {
            $downloadedPath = Download-OneFile -Url $url -DownloadDir $zipDir -ProxyRaw $ProxyRaw
            $resultText = "OK"
        }
        catch {
            $resultText = "DOWNLOAD_ERROR: " + $_.Exception.Message
        }

        Write-ProgressFile -Stage "DOWNLOAD_END" -Row $rowNo -Message "$resultText / $downloadedPath"

        if ($resultText -eq "OK") {

            Write-ProgressFile -Stage "POST_START" -Row $rowNo -Message $downloadedPath

            try {
                $post = Process-DownloadedFile `
                    -DownloadedPath $downloadedPath `
                    -ExtractRootDir $extractDir `
                    -PdfRootDir $pdfDir `
                    -HtmlRootDir $htmlDir

                $extractPath = $post.ExtractPath
                $pdfResult = $post.PdfResult
                $htmlResult = $post.HtmlResult
                $resultText = $post.ResultText
            }
            catch {
                $resultText = "POST_ERROR: " + $_.Exception.Message
            }

            Write-ProgressFile -Stage "POST_END" -Row $rowNo -Message $resultText
        }

        $resultLines.Add(
            (Clean-Field $rowNo) + "`t" +
            (Clean-Field $url) + "`t" +
            (Clean-Field $downloadedPath) + "`t" +
            (Clean-Field $extractPath) + "`t" +
            (Clean-Field $pdfResult) + "`t" +
            (Clean-Field $htmlResult) + "`t" +
            (Clean-Field $resultText)
        )
    }

    [System.IO.File]::WriteAllLines($resultPath, $resultLines, [System.Text.Encoding]::UTF8)

    if ($script:WordApp -ne $null) {
        Write-ProgressFile -Stage "WORD_QUIT" -Row "" -Message "quit"
        try {
            $script:WordApp.Quit()
        }
        catch {
        }
        $script:WordApp = $null
    }

    Set-Content -LiteralPath $donePath -Value "done" -Encoding UTF8
    Write-ProgressFile -Stage "DONE" -Row "" -Message "worker done"
}

# ============================================================
# Download
# ============================================================

function Download-OneFile {
    param(
        [string]$Url,
        [string]$DownloadDir,
        [string]$ProxyRaw
    )

    Ensure-Dir $DownloadDir

    $landingUrl = Get-LandingUrl $Url
    $cookieHeader = Get-CookieCached -LandingUrl $landingUrl -ProxyRaw $ProxyRaw

    $req = New-Object -ComObject "WinHttp.WinHttpRequest.5.1"

    Apply-Proxy -Req $req -ProxyRaw $ProxyRaw

    try {
        $req.SetTimeouts(30000, 30000, 30000, 180000)
    }
    catch {
    }

    try {
        $req.SetAutoLogonPolicy(0)
    }
    catch {
    }

    $req.Open("GET", $Url, $false)

    $req.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
    $req.SetRequestHeader("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8")
    $req.SetRequestHeader("Accept-Language", "en-US,en;q=0.9,ja;q=0.8")
    $req.SetRequestHeader("Connection", "keep-alive")
    $req.SetRequestHeader("Upgrade-Insecure-Requests", "1")

    if (-not [string]::IsNullOrWhiteSpace($cookieHeader)) {
        $req.SetRequestHeader("Cookie", $cookieHeader)
    }

    $req.Send()

    $status = [int]$req.Status

    if ($status -lt 200 -or $status -ge 300) {
        throw "HTTP $status $($req.StatusText)"
    }

    $contentDisposition = Get-HeaderSafe -Req $req -Name "Content-Disposition"
    $contentType = Get-HeaderSafe -Req $req -Name "Content-Type"

    $fileName = Get-FileNameFromContentDisposition $contentDisposition

    if ([string]::IsNullOrWhiteSpace($fileName)) {
        $fileName = Get-FileNameFromUrl $Url
    }

    if ([string]::IsNullOrWhiteSpace($fileName)) {
        $fileName = "download_" + (Get-Date -Format "yyyyMMdd_HHmmss")
    }

    $fileName = Sanitize-FileName $fileName

    if (-not [System.IO.Path]::HasExtension($fileName)) {
        $fileName += Get-ExtensionFromContentType $contentType
    }

    $savePath = Get-UniqueFilePath (Join-Path $DownloadDir $fileName)

    $bytes = [byte[]]$req.ResponseBody
    [System.IO.File]::WriteAllBytes($savePath, $bytes)

    return $savePath
}

function Get-CookieCached {
    param(
        [string]$LandingUrl,
        [string]$ProxyRaw
    )

    if ($script:CookieCache.ContainsKey($LandingUrl)) {
        return $script:CookieCache[$LandingUrl]
    }

    Write-ProgressFile -Stage "COOKIE_FETCH" -Row "" -Message $LandingUrl

    $cookie = Fetch-CookieFromLanding -LandingUrl $LandingUrl -ProxyRaw $ProxyRaw
    $script:CookieCache[$LandingUrl] = $cookie

    return $cookie
}

function Fetch-CookieFromLanding {
    param(
        [string]$LandingUrl,
        [string]$ProxyRaw
    )

    try {
        $req = New-Object -ComObject "WinHttp.WinHttpRequest.5.1"

        Apply-Proxy -Req $req -ProxyRaw $ProxyRaw

        try {
            $req.SetTimeouts(30000, 30000, 30000, 60000)
        }
        catch {
        }

        try {
            $req.SetAutoLogonPolicy(0)
        }
        catch {
        }

        $req.Open("GET", $LandingUrl, $false)
        $req.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
        $req.SetRequestHeader("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8")
        $req.SetRequestHeader("Accept-Language", "en-US,en;q=0.9,ja;q=0.8")
        $req.SetRequestHeader("Connection", "keep-alive")
        $req.Send()

        $headers = $req.GetAllResponseHeaders()
        return Parse-SetCookieHeaders $headers
    }
    catch {
        return ""
    }
}

function Parse-SetCookieHeaders {
    param([string]$Headers)

    $result = New-Object System.Collections.Generic.List[string]

    foreach ($line in ($Headers -split "`r?`n")) {
        if ($line -match "^(?i)set-cookie:\s*(.+)$") {
            $v = $Matches[1].Trim()
            $semi = $v.IndexOf(";")
            if ($semi -ge 0) {
                $v = $v.Substring(0, $semi)
            }
            if ($v.Length -gt 0) {
                $result.Add($v)
            }
        }
    }

    return ($result -join "; ")
}

function Apply-Proxy {
    param(
        $Req,
        [string]$ProxyRaw
    )

    $proxyText = Build-WinHttpProxyString $ProxyRaw

    if ([string]::IsNullOrWhiteSpace($proxyText)) {
        $Req.SetProxy(0)
    }
    else {
        $Req.SetProxy(2, $proxyText)
    }
}

function Build-WinHttpProxyString {
    param([string]$ProxyRaw)

    $p = $ProxyRaw.Trim()

    if ([string]::IsNullOrWhiteSpace($p)) {
        return ""
    }

    if ($p.Contains("=") -or $p.Contains(";")) {
        return $p
    }

    $p = $p -replace "^https?://", ""
    $p = $p.TrimEnd("/")

    return "http=$p;https=$p"
}

function Get-HeaderSafe {
    param(
        $Req,
        [string]$Name
    )

    try {
        return [string]$Req.GetResponseHeader($Name)
    }
    catch {
        return ""
    }
}

# ============================================================
# Post Process
# ============================================================

function Process-DownloadedFile {
    param(
        [string]$DownloadedPath,
        [string]$ExtractRootDir,
        [string]$PdfRootDir,
        [string]$HtmlRootDir
    )

    $baseName = Sanitize-FileName ([System.IO.Path]::GetFileNameWithoutExtension($DownloadedPath))

    $extractPath = ""
    $pdfResult = ""
    $htmlResult = ""
    $resultText = ""

    if (Test-IsZipFile $DownloadedPath) {

        $thisExtractDir = Get-UniqueFolderPath (Join-Path $ExtractRootDir $baseName)
        $thisPdfDir = Get-UniqueFolderPath (Join-Path $PdfRootDir $baseName)

        Ensure-Dir $thisExtractDir
        Ensure-Dir $thisPdfDir
        Ensure-Dir $HtmlRootDir

        Write-ProgressFile -Stage "EXTRACT_START" -Row "" -Message $DownloadedPath

        Expand-Archive -LiteralPath $DownloadedPath -DestinationPath $thisExtractDir -Force

        Write-ProgressFile -Stage "EXTRACT_END" -Row "" -Message $thisExtractDir

        $pdfCount = Convert-FolderFilesToPdf -SourceFolder $thisExtractDir -PdfFolder $thisPdfDir
        $htmlCount = Convert-FolderWordFilesToBeautifulHtml -SourceFolder $thisExtractDir -HtmlFolder $HtmlRootDir

        $extractPath = $thisExtractDir
        $pdfResult = "$pdfCount 件 PDF作成"
        $htmlResult = "$htmlCount 件 HTML作成"
        $resultText = "ZIP展開OK / PDF作成OK / HTML作成OK"
    }
    else {

        Ensure-Dir $PdfRootDir
        Ensure-Dir $HtmlRootDir

        $onePdf = Convert-SingleFileToPdf -FilePath $DownloadedPath -PdfFolder $PdfRootDir

        if (-not [string]::IsNullOrWhiteSpace($onePdf)) {
            $pdfResult = $onePdf
        }
        else {
            $pdfResult = "PDF作成対象外"
        }

        $oneHtml = Convert-SingleWordFileToBeautifulHtml -FilePath $DownloadedPath -HtmlFolder $HtmlRootDir

        if (-not [string]::IsNullOrWhiteSpace($oneHtml)) {
            $htmlResult = $oneHtml
        }
        else {
            $htmlResult = "HTML作成対象外"
        }

        $resultText = "PDF/HTML作成処理完了"
    }

    return [pscustomobject]@{
        ExtractPath = $extractPath
        PdfResult = $pdfResult
        HtmlResult = $htmlResult
        ResultText = $resultText
    }
}

# ============================================================
# PDF
# ============================================================

function Convert-FolderFilesToPdf {
    param(
        [string]$SourceFolder,
        [string]$PdfFolder
    )

    if (-not (Test-Path -LiteralPath $SourceFolder)) {
        return 0
    }

    Ensure-Dir $PdfFolder

    $count = 0

    $files = Get-ChildItem -LiteralPath $SourceFolder -File -Recurse -ErrorAction SilentlyContinue

    foreach ($file in $files) {
        $pdfPath = Convert-SingleFileToPdf -FilePath $file.FullName -PdfFolder $PdfFolder
        if (-not [string]::IsNullOrWhiteSpace($pdfPath)) {
            $count++
        }
    }

    return $count
}

function Convert-SingleFileToPdf {
    param(
        [string]$FilePath,
        [string]$PdfFolder
    )

    if (-not (Test-Path -LiteralPath $FilePath)) {
        return ""
    }

    Ensure-Dir $PdfFolder

    $ext = [System.IO.Path]::GetExtension($FilePath).TrimStart(".").ToLower()
    $base = Sanitize-FileName ([System.IO.Path]::GetFileNameWithoutExtension($FilePath))
    $pdfPath = Get-UniqueFilePath (Join-Path $PdfFolder ($base + ".pdf"))

    try {
        switch ($ext) {

            "pdf" {
                Copy-Item -LiteralPath $FilePath -Destination $pdfPath -Force
                return $pdfPath
            }

            { $_ -in @("html", "htm", "txt", "rtf", "doc", "docx", "docm") } {
                if (-not (Ensure-WordApp)) {
                    return ""
                }
                Convert-WordOpenableFileToPdf -FilePath $FilePath -PdfPath $pdfPath
                if (Test-Path -LiteralPath $pdfPath) {
                    return $pdfPath
                }
                return ""
            }

            { $_ -in @("jpg", "jpeg", "png", "bmp", "gif", "tif", "tiff") } {
                if (-not (Ensure-WordApp)) {
                    return ""
                }
                Convert-ImageToPdf -ImagePath $FilePath -PdfPath $pdfPath
                if (Test-Path -LiteralPath $pdfPath) {
                    return $pdfPath
                }
                return ""
            }

            default {
                return ""
            }
        }
    }
    catch {
        return ""
    }
}

function Convert-WordOpenableFileToPdf {
    param(
        [string]$FilePath,
        [string]$PdfPath
    )

    $doc = $null

    try {
        $doc = $script:WordApp.Documents.Open($FilePath, $false, $true, $false)
        $doc.ExportAsFixedFormat($PdfPath, 17)
    }
    finally {
        if ($doc -ne $null) {
            $doc.Close($false)
        }
    }
}

function Convert-ImageToPdf {
    param(
        [string]$ImagePath,
        [string]$PdfPath
    )

    $doc = $null

    try {
        $doc = $script:WordApp.Documents.Add()

        $doc.PageSetup.TopMargin = 36
        $doc.PageSetup.BottomMargin = 36
        $doc.PageSetup.LeftMargin = 36
        $doc.PageSetup.RightMargin = 36

        $availableWidth = $doc.PageSetup.PageWidth - $doc.PageSetup.LeftMargin - $doc.PageSetup.RightMargin
        $availableHeight = $doc.PageSetup.PageHeight - $doc.PageSetup.TopMargin - $doc.PageSetup.BottomMargin

        $img = $doc.InlineShapes.AddPicture($ImagePath, $false, $true)
        $img.LockAspectRatio = $true

        if ($img.Width -gt $availableWidth) {
            $img.Width = $availableWidth
        }

        if ($img.Height -gt $availableHeight) {
            $img.Height = $availableHeight
        }

        $doc.ExportAsFixedFormat($PdfPath, 17)
    }
    finally {
        if ($doc -ne $null) {
            $doc.Close($false)
        }
    }
}

# ============================================================
# HTML
# ============================================================

function Convert-FolderWordFilesToBeautifulHtml {
    param(
        [string]$SourceFolder,
        [string]$HtmlFolder
    )

    if (-not (Test-Path -LiteralPath $SourceFolder)) {
        return 0
    }

    Ensure-Dir $HtmlFolder

    $count = 0
    $files = Get-ChildItem -LiteralPath $SourceFolder -File -Recurse -ErrorAction SilentlyContinue

    foreach ($file in $files) {
        $htmlPath = Convert-SingleWordFileToBeautifulHtml -FilePath $file.FullName -HtmlFolder $HtmlFolder
        if (-not [string]::IsNullOrWhiteSpace($htmlPath)) {
            $count++
        }
    }

    return $count
}

function Convert-SingleWordFileToBeautifulHtml {
    param(
        [string]$FilePath,
        [string]$HtmlFolder
    )

    if (-not (Test-Path -LiteralPath $FilePath)) {
        return ""
    }

    $ext = [System.IO.Path]::GetExtension($FilePath).TrimStart(".").ToLower()

    if ($ext -notin @("doc", "docx", "docm", "rtf", "txt", "html", "htm")) {
        return ""
    }

    Ensure-Dir $HtmlFolder

    $baseName = Sanitize-FileName ([System.IO.Path]::GetFileNameWithoutExtension($FilePath))
    $outHtmlPath = Get-UniqueFilePath (Join-Path $HtmlFolder ($baseName + ".html"))

    try {
        if ($ext -in @("doc", "docx", "docm", "rtf", "txt")) {

            if (-not (Ensure-WordApp)) {
                return ""
            }

            Save-WordOpenableFileAsFilteredHtml -FilePath $FilePath -OutHtmlPath $outHtmlPath
            $rawHtml = Read-TextUtf8 $outHtmlPath
        }
        else {
            $rawHtml = Read-TextUtf8 $FilePath
        }

        $bodyHtml = Extract-BodyInnerHtml $rawHtml

        if ([string]::IsNullOrWhiteSpace($bodyHtml)) {
            $bodyHtml = "<pre>" + (Html-Escape (Read-TextDefault $FilePath)) + "</pre>"
        }

        $prettyHtml = Build-BeautifulHtmlDocument `
            -TitleText $baseName `
            -SourceFileName ([System.IO.Path]::GetFileName($FilePath)) `
            -BodyHtml $bodyHtml

        Write-TextUtf8 -Path $outHtmlPath -Text $prettyHtml

        if (Test-Path -LiteralPath $outHtmlPath) {
            return $outHtmlPath
        }

        return ""
    }
    catch {
        return ""
    }
}

function Save-WordOpenableFileAsFilteredHtml {
    param(
        [string]$FilePath,
        [string]$OutHtmlPath
    )

    $doc = $null

    try {
        $doc = $script:WordApp.Documents.Open($FilePath, $false, $true, $false)

        try {
            $doc.WebOptions.Encoding = 65001
        }
        catch {
        }

        $doc.SaveAs2($OutHtmlPath, 10)
    }
    finally {
        if ($doc -ne $null) {
            $doc.Close($false)
        }
    }
}

function Build-BeautifulHtmlDocument {
    param(
        [string]$TitleText,
        [string]$SourceFileName,
        [string]$BodyHtml
    )

    $titleEsc = Html-Escape $TitleText
    $sourceEsc = Html-Escape $SourceFileName

    return @"
<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$titleEsc</title>
$(Beautiful-HtmlCss)
</head>
<body>
<main class="page">
<header class="doc-header">
<div class="label">WORD HTML VIEW</div>
<h1>$titleEsc</h1>
<div class="meta">Source: $sourceEsc</div>
</header>
<article class="doc-body">
$BodyHtml
</article>
</main>
</body>
</html>
"@
}

function Beautiful-HtmlCss {
    return @"
<style>
html{background:#f3f4f6;color:#111827;}
body{margin:0;font-family:-apple-system,BlinkMacSystemFont,'Yu Gothic','Meiryo','Segoe UI',sans-serif;line-height:1.85;}
.page{max-width:980px;margin:32px auto;padding:40px;background:#fff;border-radius:18px;box-shadow:0 12px 36px rgba(15,23,42,.12);}
.doc-header{border-bottom:1px solid #e5e7eb;margin-bottom:28px;padding-bottom:18px;}
.label{font-size:12px;letter-spacing:.14em;color:#6b7280;font-weight:700;margin-bottom:8px;}
h1{font-size:26px;line-height:1.35;margin:0 0 8px 0;color:#0f172a;}
.meta{font-size:13px;color:#6b7280;}
.doc-body{font-size:15.5px;}
.doc-body p{margin:.65em 0;}
.doc-body h1,.doc-body h2,.doc-body h3{line-height:1.45;margin-top:1.4em;color:#111827;}
.doc-body table{border-collapse:collapse;width:100%;margin:18px 0;font-size:14px;}
.doc-body th,.doc-body td{border:1px solid #d1d5db;padding:8px 10px;vertical-align:top;}
.doc-body th{background:#f9fafb;font-weight:700;}
.doc-body img{max-width:100%;height:auto;display:block;margin:16px auto;}
.doc-body pre{white-space:pre-wrap;background:#f9fafb;border:1px solid #e5e7eb;padding:16px;border-radius:12px;}
.doc-body a{color:#2563eb;text-decoration:none;}
.doc-body a:hover{text-decoration:underline;}
@media print{html{background:#fff}.page{margin:0;box-shadow:none;border-radius:0}}
</style>
"@
}

# ============================================================
# HTML Combine
# ============================================================

function Combine-HtmlFilesByFive {
    param(
        [string]$HtmlDir,
        [string]$OutputDir
    )

    Ensure-Dir $OutputDir

    if (-not (Test-Path -LiteralPath $HtmlDir)) {
        return 0
    }

    $files = @(Get-ChildItem -LiteralPath $HtmlDir -File -Filter "*.html" -ErrorAction SilentlyContinue | Sort-Object Name)

    if ($files.Count -eq 0) {
        return 0
    }

    $groupSize = 5
    $groupNo = 0

    for ($i = 0; $i -lt $files.Count; $i += $groupSize) {

        $groupNo++
        $end = [Math]::Min($i + $groupSize - 1, $files.Count - 1)

        $sections = New-Object System.Collections.Generic.List[string]

        for ($j = $i; $j -le $end; $j++) {

            $oneFile = $files[$j]
            $raw = Read-TextUtf8 $oneFile.FullName
            $body = Extract-BodyInnerHtml $raw
            $body = Prefix-RelativeLinksForCombinedHtml -Html $body -Prefix "../HTML/"

            $sections.Add(@"
<section class="combined-doc">
<div class="source-title">$([System.Net.WebUtility]::HtmlEncode($oneFile.Name))</div>
$body
</section>
"@)
        }

        $title = "HTML_COMBINE_{0:000}" -f $groupNo
        $outPath = Join-Path $OutputDir ($title + ".html")
        $outPath = Get-UniqueFilePath $outPath

        $html = @"
<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$title</title>
$(Combined-HtmlCss)
</head>
<body>
<main class="page">
<header class="combine-header">
<div class="label">HTML COMBINE</div>
<h1>$title</h1>
<div class="meta">$($i + 1)件目〜$($end + 1)件目 / 全$($files.Count)件</div>
</header>
$($sections -join "`r`n")
</main>
</body>
</html>
"@

        Write-TextUtf8 -Path $outPath -Text $html
    }

    return $groupNo
}

function Combined-HtmlCss {
    return @"
<style>
html{background:#eef2f7;color:#111827;}
body{margin:0;font-family:-apple-system,BlinkMacSystemFont,'Yu Gothic','Meiryo','Segoe UI',sans-serif;line-height:1.85;}
.page{max-width:1080px;margin:32px auto;padding:40px;background:#fff;border-radius:18px;box-shadow:0 12px 36px rgba(15,23,42,.12);}
.combine-header{border-bottom:2px solid #d1d5db;margin-bottom:28px;padding-bottom:18px;}
.label{font-size:12px;letter-spacing:.14em;color:#6b7280;font-weight:700;margin-bottom:8px;}
h1{font-size:28px;line-height:1.35;margin:0 0 8px 0;color:#0f172a;}
.meta{font-size:13px;color:#6b7280;}
.combined-doc{padding:28px 0;border-top:1px dashed #cbd5e1;}
.combined-doc:first-of-type{border-top:none;}
.source-title{position:sticky;top:0;background:#111827;color:#fff;padding:8px 12px;border-radius:10px;font-size:13px;font-weight:700;margin-bottom:16px;z-index:1;}
p{margin:.65em 0;}
h1,h2,h3{line-height:1.45;margin-top:1.4em;color:#111827;}
table{border-collapse:collapse;width:100%;margin:18px 0;font-size:14px;}
th,td{border:1px solid #d1d5db;padding:8px 10px;vertical-align:top;}
th{background:#f9fafb;font-weight:700;}
img{max-width:100%;height:auto;display:block;margin:16px auto;}
pre{white-space:pre-wrap;background:#f9fafb;border:1px solid #e5e7eb;padding:16px;border-radius:12px;}
a{color:#2563eb;text-decoration:none;}
a:hover{text-decoration:underline;}
@media print{html{background:#fff}.page{margin:0;box-shadow:none;border-radius:0}.source-title{position:static}}
</style>
"@
}

function Prefix-RelativeLinksForCombinedHtml {
    param(
        [string]$Html,
        [string]$Prefix
    )

    $pattern = '(?i)\b(src|href)\s*=\s*("([^"]*)"|''([^'']*)'')'

    return [regex]::Replace($Html, $pattern, {
        param($m)

        $attr = $m.Groups[1].Value
        $quoteAndValue = $m.Groups[2].Value
        $url = $m.Groups[3].Value

        if ([string]::IsNullOrEmpty($url)) {
            $url = $m.Groups[4].Value
        }

        if (Should-PrefixRelativeUrl $url) {
            $quote = $quoteAndValue.Substring(0, 1)
            return "$attr=$quote$Prefix$url$quote"
        }

        return $m.Value
    })
}

function Should-PrefixRelativeUrl {
    param([string]$Url)

    $u = $Url.Trim().ToLower()

    if ($u.Length -eq 0) { return $false }
    if ($u.StartsWith("#")) { return $false }
    if ($u.StartsWith("http:")) { return $false }
    if ($u.StartsWith("https:")) { return $false }
    if ($u.StartsWith("data:")) { return $false }
    if ($u.StartsWith("mailto:")) { return $false }
    if ($u.StartsWith("tel:")) { return $false }
    if ($u.StartsWith("../")) { return $false }
    if ($u.StartsWith("/")) { return $false }
    if ($u -match "^[a-z]:\\") { return $false }

    return $true
}

# ============================================================
# Word
# ============================================================

function Ensure-WordApp {

    if ($script:WordApp -ne $null) {
        return $true
    }

    try {
        Write-ProgressFile -Stage "WORD_START" -Row "" -Message "Microsoft Word 起動中"

        $script:WordApp = New-Object -ComObject "Word.Application"
        $script:WordApp.Visible = $false
        $script:WordApp.DisplayAlerts = 0

        Write-ProgressFile -Stage "WORD_READY" -Row "" -Message "Microsoft Word 起動完了"

        return $true
    }
    catch {
        Write-ProgressFile -Stage "WORD_ERROR" -Row "" -Message $_.Exception.Message
        return $false
    }
}

# ============================================================
# URL / filename
# ============================================================

function Get-LandingUrl {
    param([string]$Url)

    try {
        $uri = [Uri]$Url
        return $uri.Scheme + "://" + $uri.Authority + "/"
    }
    catch {
        return $Url
    }
}

function Get-FileNameFromUrl {
    param([string]$Url)

    try {
        $uri = [Uri]$Url
        $name = [System.IO.Path]::GetFileName($uri.AbsolutePath)
        return [System.Uri]::UnescapeDataString($name)
    }
    catch {
        return ""
    }
}

function Get-FileNameFromContentDisposition {
    param([string]$ContentDisposition)

    if ([string]::IsNullOrWhiteSpace($ContentDisposition)) {
        return ""
    }

    if ($ContentDisposition -match "(?i)filename\*\s*=\s*([^;]+)") {
        $v = $Matches[1].Trim().Trim('"')
        $v = $v -replace "^UTF-8''", ""
        return [System.Uri]::UnescapeDataString($v)
    }

    if ($ContentDisposition -match "(?i)filename\s*=\s*([^;]+)") {
        return $Matches[1].Trim().Trim('"')
    }

    return ""
}

function Get-ExtensionFromContentType {
    param([string]$ContentType)

    $ct = $ContentType.ToLower()

    if ($ct.Contains("zip")) { return ".zip" }
    if ($ct.Contains("pdf")) { return ".pdf" }
    if ($ct.Contains("html")) { return ".html" }
    if ($ct.Contains("xml")) { return ".xml" }
    if ($ct.Contains("json")) { return ".json" }
    if ($ct.Contains("plain")) { return ".txt" }
    if ($ct.Contains("png")) { return ".png" }
    if ($ct.Contains("jpeg") -or $ct.Contains("jpg")) { return ".jpg" }

    return ".bin"
}

function Sanitize-FileName {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return "page"
    }

    $invalid = [System.IO.Path]::GetInvalidFileNameChars()

    foreach ($ch in $invalid) {
        $Name = $Name.Replace($ch, "_")
    }

    $Name = $Name.Trim().TrimEnd(".")

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return "page"
    }

    return $Name
}

# ============================================================
# File helpers
# ============================================================

function Normalize-FolderPath {
    param([string]$Path)

    if ($null -eq $Path) {
        return ""
    }

    $p = $Path.Trim().Trim('"')

    while ($p.Length -gt 3 -and ($p.EndsWith("\") -or $p.EndsWith("/"))) {
        $p = $p.Substring(0, $p.Length - 1)
    }

    return $p
}

function Ensure-Dir {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Get-UniqueFilePath {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return $Path
    }

    $dir = [System.IO.Path]::GetDirectoryName($Path)
    $base = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $ext = [System.IO.Path]::GetExtension($Path)

    for ($i = 1; $i -le 9999; $i++) {
        $candidate = Join-Path $dir ("{0}_{1:000}{2}" -f $base, $i, $ext)
        if (-not (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }

    return $Path
}

function Get-UniqueFolderPath {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return $Path
    }

    for ($i = 1; $i -le 9999; $i++) {
        $candidate = "{0}_{1:000}" -f $Path, $i
        if (-not (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }

    return $Path
}

function Test-IsZipFile {
    param([string]$Path)

    $ext = [System.IO.Path]::GetExtension($Path).ToLower()

    if ($ext -eq ".zip") {
        return $true
    }

    try {
        $fs = [System.IO.File]::OpenRead($Path)
        $b1 = $fs.ReadByte()
        $b2 = $fs.ReadByte()
        $fs.Close()

        return ($b1 -eq 80 -and $b2 -eq 75)
    }
    catch {
        return $false
    }
}

function Read-TextUtf8 {
    param([string]$Path)

    try {
        return [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
    }
    catch {
        return Read-TextDefault $Path
    }
}

function Read-TextDefault {
    param([string]$Path)

    try {
        return [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::Default)
    }
    catch {
        return ""
    }
}

function Write-TextUtf8 {
    param(
        [string]$Path,
        [string]$Text
    )

    [System.IO.File]::WriteAllText($Path, $Text, [System.Text.Encoding]::UTF8)
}

function Extract-BodyInnerHtml {
    param([string]$Html)

    if ([string]::IsNullOrWhiteSpace($Html)) {
        return ""
    }

    $m = [regex]::Match($Html, "(?is)<body[^>]*>(.*?)</body>")

    if ($m.Success) {
        return $m.Groups[1].Value
    }

    return $Html
}

function Html-Escape {
    param([string]$Text)

    return [System.Net.WebUtility]::HtmlEncode($Text)
}

function Clean-Field {
    param([string]$Text)

    if ($null -eq $Text) {
        return ""
    }

    return ($Text -replace "`t", " " -replace "`r", " " -replace "`n", " ")
}

function Write-ProgressFile {
    param(
        [string]$Stage,
        [string]$Row,
        [string]$Message
    )

    if ([string]::IsNullOrWhiteSpace($script:ProgressPath)) {
        return
    }

    $text = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tstage=$Stage`trow=$Row`t$Message"

    try {
        Set-Content -LiteralPath $script:ProgressPath -Value $text -Encoding UTF8
    }
    catch {
    }
}