<#
.SYNOPSIS
  verify_stats.ps1 — генерация N копий отчёта, извлечение таблицы статистики
  и сравнение SHA-256 хешей.

.ПАРАМЕТРЫ
  -SheetName   Лист данных (pass / fail). По умолчанию pass.
  -Runs        Количество запусков. По умолчанию 3.
  -Mode        Режим (draft / final). По умолчанию final.

.ПРИМЕРЫ
  # 3 запуска, лист pass (по умолчанию)
  Powershell -ExecutionPolicy Bypass .\verify_stats.ps1

  # 5 запусков, лист fail
  Powershell -ExecutionPolicy Bypass .\verify_stats.ps1 -SheetName fail -Runs 5
#>

param(
    [ValidateSet("pass","fail")]
    [string]$SheetName = "pass",
    [int]$Runs = 3,
    [ValidateSet("draft","final")]
    [string]$Mode = "final"
)

Add-Type -AssemblyName System.IO.Compression.FileSystem

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ScriptPath = Join-Path $ScriptDir "mvp.ps1"

Write-Host "=== verify_stats: $Runs запусков, лист '$SheetName' ===" -ForegroundColor Magenta

$hashes = @()
for ($i = 1; $i -le $Runs; $i++) {
    Get-Process excel, winword -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 2

    $outName = "Report_${SheetName}_$i.docx"
    $outPath = Join-Path $ScriptDir $outName

    # Генерация
    & $ScriptPath -Mode $Mode -sheetName $SheetName -ReportPath $outName | Out-Null
    Get-Process excel, winword -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 1

    if (-not (Test-Path $outPath)) {
        Write-Host "ERROR: file not found: $outPath" -ForegroundColor Red
        exit 1
    }

    # Извлечение таблицы статистики (вторая таблица в документе)
    $zip = [System.IO.Compression.ZipFile]::OpenRead($outPath)
    $entry = $zip.GetEntry("word/document.xml")
    $reader = New-Object System.IO.StreamReader($entry.Open())
    $xmlText = $reader.ReadToEnd()
    $reader.Close()
    $zip.Dispose()

    $xml = [xml]$xmlText
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    $tables = $xml.SelectNodes("//w:tbl", $ns)
    $statsTable = $tables[1]

    $rows = $statsTable.SelectNodes("w:tr", $ns)
    $lines = @()
    foreach ($row in $rows) {
        $cells = $row.SelectNodes("w:tc", $ns)
        $cellTexts = @()
        foreach ($cell in $cells) {
            $texts = $cell.SelectNodes(".//w:t", $ns)
            $txt = ($texts | ForEach-Object { $_.InnerText }) -join ""
            $cellTexts += $txt
        }
        if ($cellTexts.Count -ge 2) {
            $lines += $cellTexts -join " | "
        }
    }
    $tableContent = $lines -join "`n"
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($tableContent)
    $hash = [System.Security.Cryptography.SHA256]::Create().ComputeHash($bytes)
    $hashStr = [System.BitConverter]::ToString($hash) -replace '-'

    $hashes += @{ Run = $i; File = $outName; Hash = $hashStr; Content = $lines }
}

# Вывод
Write-Host "`nРезультаты:" -ForegroundColor Cyan
$hashes | ForEach-Object {
    Write-Host "  #$($_.Run)  $($_.File)" -ForegroundColor Yellow
    $_.Content | ForEach-Object { Write-Host "     $_" }
    Write-Host "     SHA-256: $($_.Hash)" -ForegroundColor Green
    Write-Host ""
}

# Сравнение
$ref = $hashes[0].Hash
$allMatch = $true
$hashes | Select-Object -Skip 1 | ForEach-Object {
    if ($_.Hash -ne $ref) { $allMatch = $false }
}

if ($allMatch) {
    Write-Host "OK: все $Runs SHA-256 хешей совпадают: $ref" -ForegroundColor Green
} else {
    Write-Host "FAIL: хеши различаются!" -ForegroundColor Red
}
