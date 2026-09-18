# PrintCalc - Adobe Acrobat 文档转 PDF
# 通过 Acrobat 的 IAC (AcroExch.AVDoc / PDDoc) 把任意受支持文档转换为 PDF。
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
        if ($line.Length -gt 0) {
            $items += ($line | ConvertFrom-Json)
        }
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
$acro = $null
try {
    $acro = New-Object -ComObject AcroExch.App
    try { $acro.Hide() } catch { }
}
catch {
    foreach ($j in $jobs) {
        Write-JsonLine $ResultFile ([pscustomobject]@{ src = $j.src; dst = $j.dst; ok = $false; pages = 0; error = "无法启动 Adobe Acrobat：$($_.Exception.Message)" })
    }
    exit 1
}

foreach ($job in $jobs) {
    $src = $job.src
    $dst = $job.dst
    $pages = 0
    $ok = $false
    $errorText = ""
    try {
        $ext = [System.IO.Path]::GetExtension($src).ToLowerInvariant()
        if ($ext -eq '.pdf') {
            Copy-Item -LiteralPath $src -Destination $dst -Force
            $ok = $true
        }
        else {
            $avdoc = New-Object -ComObject AcroExch.AVDoc
            $opened = $avdoc.Open($src, "")
            if (-not $opened) {
                throw "Acrobat 无法打开该文件（格式不受支持或文件损坏）"
            }
            $pddoc = $avdoc.GetPDDoc()
            if ($pddoc -ne $null) { $pages = $pddoc.GetNumPages() }
            $saved = $pddoc.Save(1, $dst)
            try { $avdoc.Close($true) | Out-Null } catch { }
            if (-not $saved) { throw "Acrobat 保存 PDF 失败" }
            $ok = $true
        }
    }
    catch {
        $errorText = $_.Exception.Message
    }
    Write-JsonLine $ResultFile ([pscustomobject]@{ src = $src; dst = $dst; ok = $ok; pages = $pages; error = $errorText })
}

try { $acro.Exit() } catch { }
exit 0
