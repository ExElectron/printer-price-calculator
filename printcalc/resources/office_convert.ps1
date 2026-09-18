# PrintCalc - Microsoft Office 文档转 PDF（Acrobat 失败时的备选方案）
# 入参：-JobsFile  JSONL，每行 {"src":"...","dst":"..."}
# 出参：-ResultFile JSONL，每行 {"src":..,"dst":..,"ok":true/false,"pages":n,"error":".."}
param(
    [Parameter(Mandatory = $true)][string]$JobsFile,
    [Parameter(Mandatory = $true)][string]$ResultFile
)

$ErrorActionPreference = 'Continue'

function Read-JsonLines([string]$path) {
    $items = @()
    foreach ($line in [System.IO.File]::ReadAllLines($path, [System.Text.Encoding]::UTF8)) {
        $line = $line.Trim()
        if ($line.Length -gt 0) { $items += ($line | ConvertFrom-Json) }
    }
    return $items
}

function Write-JsonLine([string]$path, $object) {
    $json = $object | ConvertTo-Json -Compress -Depth 5
    [System.IO.File]::AppendAllText($path, $json + "`n", (New-Object System.Text.UTF8Encoding($false)))
}

if (Test-Path -LiteralPath $ResultFile) { Remove-Item -LiteralPath $ResultFile -Force }
[System.IO.File]::WriteAllText($ResultFile, "", (New-Object System.Text.UTF8Encoding($false)))

$jobs = Read-JsonLines $JobsFile

foreach ($job in $jobs) {
    $src = $job.src
    $dst = $job.dst
    $ok = $false
    $errorText = ""
    $app = $null
    $doc = $null
    try {
        $ext = [System.IO.Path]::GetExtension($src).ToLowerInvariant()
        switch -Regex ($ext) {
            '^\.(doc|docx|rtf|odt|txt)$' {
                $app = New-Object -ComObject Word.Application
                $app.Visible = $false
                $app.DisplayAlerts = 0
                $doc = $app.Documents.Open($src, $false, $true)
                $doc.ExportAsFixedFormat($dst, 17)
                $doc.Close(0)
                $app.Quit()
                $ok = $true
            }
            '^\.(xls|xlsx|xlsm|csv|ods)$' {
                $app = New-Object -ComObject Excel.Application
                $app.Visible = $false
                $app.DisplayAlerts = $false
                $doc = $app.Workbooks.Open($src)
                $doc.ExportAsFixedFormat(0, $dst)
                $doc.Close($false)
                $app.Quit()
                $ok = $true
            }
            '^\.(ppt|pptx|odp)$' {
                $app = New-Object -ComObject PowerPoint.Application
                $doc = $app.Presentations.Open($src, $true, $false, $false)
                $doc.ExportAsFixedFormat($dst, 2)
                $doc.Close()
                $app.Quit()
                $ok = $true
            }
            default {
                throw "Office 引擎不支持该格式：$ext"
            }
        }
    }
    catch {
        $errorText = $_.Exception.Message
        try { if ($doc -ne $null) { $doc.Close(0) } } catch { }
        try { if ($app -ne $null) { $app.Quit() } } catch { }
    }
    Write-JsonLine $ResultFile ([pscustomobject]@{ src = $src; dst = $dst; ok = $ok; pages = 0; error = $errorText })
}

exit 0
