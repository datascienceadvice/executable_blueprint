param(
    [ValidateSet("draft","final")]
    [string]$Mode = "final",

    [string]$DataPath = "data.xlsx",

    [string]$sheetName = "pass",

    [string]$ReportPath = ""
)

if (-not $ReportPath) {
    $suffix = if ($Mode -eq "draft") { "draft" } else { $sheetName }
    $ReportPath = "Report_$suffix.docx"
}

# Очистка предыдущих процессов Excel/Word
Get-Process excel, winword -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 1

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir
Write-Host "1. Папка скрипта: $ScriptDir" -ForegroundColor Gray

$ExcelPath = Join-Path $ScriptDir $DataPath
$OutputDocx = Join-Path $ScriptDir $ReportPath
$PercentWidth = 80   # ширина графика в % от ширины страницы

# --- Чтение данных, расчёт статистики (только для final) ---
if ($Mode -eq "final") {
    Write-Host "РЕЖИМ: ОТЧЁТ (final)" -ForegroundColor Green
    Write-Host "2. Проверка файла Excel..." -ForegroundColor Cyan
    if (-not (Test-Path $ExcelPath)) { Write-Error "Файл $ExcelPath не найден"; exit 1 }
    
    Import-Module ImportExcel -ErrorAction Stop
    $data = Import-Excel -Path $ExcelPath -WorksheetName $sheetName
    if (($data.ID -eq $null) -or ($data.area -eq $null)) { Write-Error "Файл должен содержать колонки 'ID' и 'area'"; exit 1 }

    Write-Host "3. Извлечение числовых значений..." -ForegroundColor Cyan
    $areas = @()
    foreach ($val in $data.area) {
        $num = $null
        if ($val -is [double] -or $val -is [int] -or $val -is [decimal]) { $num = [double]$val }
        elseif ($val -is [string]) { [double]::TryParse($val, [ref]$num) | Out-Null }
        if ($num -ne $null) { $areas += $num }
    }
    if ($areas.Count -eq 0) { Write-Error "Нет числовых данных в колонке area"; exit 1 }

    $count = $areas.Count
    $mean = ($areas | Measure-Object -Average).Average
    $sumSq = 0; foreach ($x in $areas) { $sumSq += ($x - $mean) * ($x - $mean) }
    $stddev = if ($count -gt 1) { [math]::Sqrt($sumSq / ($count - 1)) } else { 0 }
    $min = ($areas | Measure-Object -Minimum).Minimum
    $max = ($areas | Measure-Object -Maximum).Maximum
    $rsd = if ($mean -ne 0) { ($stddev / $mean) * 100 } else { 0 }
    Write-Host "   Статистика: Среднее=$([math]::Round($mean,4)), СКО=$([math]::Round($stddev,4)), ОСКО=$([math]::Round($rsd,4))%" -ForegroundColor Gray
} else {
    Write-Host "РЕЖИМ: ЧЕРНОВИК (draft) — используются маски" -ForegroundColor Yellow
    $maskRows = 6
}

# --- Генерация графика (через временный Excel) ---
Write-Host "4. Генерация графика..." -ForegroundColor Cyan
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    if ($Mode -eq "final") {
        # Создаём временный Excel и заполняем реальными данными
        $wb = $excel.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)
        $ws.Cells.Item(1,1) = "ID"; $ws.Cells.Item(1,2) = "Площадь"
        for ($i = 0; $i -lt $data.Count; $i++) {
            $ws.Cells.Item($i+2, 1) = $data[$i].ID
            $ws.Cells.Item($i+2, 2) = $data[$i].area
        }
        $chartTitle = "Площади пиков по образцам"
    } else {
        # Черновик: создаём фиктивные данные (нули) для образца графика
        $wb = $excel.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)
        $ws.Cells.Item(1,1) = "ID"; $ws.Cells.Item(1,2) = "Площадь"
        for ($i = 1; $i -le $maskRows; $i++) {
            $ws.Cells.Item($i+1, 1) = $i
            $ws.Cells.Item($i+1, 2) = 0
        }
        $chartTitle = "ОБРАЗЕЦ ГРАФИКА"
    }

    # Строим диаграмму
    $range = $ws.Range("A1:B$($ws.UsedRange.Rows.Count)")
    $chartObject = $ws.Shapes.AddChart()
    $chartObject.Width = 500
    $chartObject.Height = 300
    $chart = $chartObject.Chart
    $chart.ChartType = 51  # Гистограмма
    $chart.SetSourceData($range)
    $chart.HasTitle = $true
    $chart.ChartTitle.Text = $chartTitle
    $chart.ChartTitle.Font.Size = 12
    $chart.ChartTitle.Font.Bold = $true

    # Копируем диаграмму в буфер обмена
    $chart.ChartArea.Copy()
    Start-Sleep -Milliseconds 800

    # --- Создание Word-документа и вставка содержимого ---
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Add()
    $selection = $word.Selection

    # Заголовок
    $selection.Style = "Заголовок 1"
    if ($Mode -eq "draft") {
        $selection.TypeText("ПЛАН ОТЧЁТА")
    } else {
        $selection.TypeText("Отчёт о статистическом анализе хроматографических данных")
    }
    $selection.TypeParagraph()
    $selection.Style = "Обычный"

    # Источник и дата
    $selection.Font.Bold = $true
    $selection.TypeText("Источник: ")
    $selection.Font.Bold = $false
    if ($Mode -eq "draft") {
        $selection.TypeText("[данные будут получены после эксперимента]")
    } else {
        $selection.TypeText("$DataPath, лист: $sheetName")
    }
    $selection.TypeParagraph()

    $selection.Font.Bold = $true
    $selection.TypeText("Дата генерации: ")
    $selection.Font.Bold = $false
    $selection.TypeText("$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $selection.TypeParagraph()

    # Таблица исходных данных
    $selection.Style = "Заголовок 2"
    if ($Mode -eq "draft") {
        $selection.TypeText("Ожидаемая структура исходных данных (маски)")
    } else {
        $selection.TypeText("Исходные данные")
    }
    $selection.TypeParagraph()
    $selection.Style = "Обычный"

    if ($Mode -eq "draft") {
        $table1 = $doc.Tables.Add($selection.Range, $maskRows + 1, 2)
        $table1.Borders.Enable = 1
        for ($row = 1; $row -le $table1.Rows.Count; $row++) {
            for ($col = 1; $col -le $table1.Columns.Count; $col++) {
                $table1.Cell($row, $col).Range.ParagraphFormat.Alignment = 1
            }
        }
        $table1.Cell(1,1).Range.Text = "ID"; $table1.Cell(1,1).Range.Font.Bold = $true
        $table1.Cell(1,2).Range.Text = "Площадь"; $table1.Cell(1,2).Range.Font.Bold = $true
        for ($i = 1; $i -le $maskRows; $i++) {
            $table1.Cell($i + 1, 1).Range.Text = $i.ToString()
            $table1.Cell($i + 1, 2).Range.Text = "[x.xxx]"
        }
    } else {
        $table1 = $doc.Tables.Add($selection.Range, $data.Count + 1, 2)
        $table1.Borders.Enable = 1
        for ($row = 1; $row -le $table1.Rows.Count; $row++) {
            for ($col = 1; $col -le $table1.Columns.Count; $col++) {
                $table1.Cell($row, $col).Range.ParagraphFormat.Alignment = 1
            }
        }
        $table1.Cell(1,1).Range.Text = "ID"; $table1.Cell(1,1).Range.Font.Bold = $true
        $table1.Cell(1,2).Range.Text = "Площадь"; $table1.Cell(1,2).Range.Font.Bold = $true
        for ($i = 0; $i -lt $data.Count; $i++) {
            $table1.Cell($i + 2, 1).Range.Text = $data[$i].ID.ToString()
            $table1.Cell($i + 2, 2).Range.Text = $data[$i].area.ToString()
        }
    }
    $selection.EndKey(6) | Out-Null
    #$selection.TypeParagraph()

    # Таблица статистики
    $selection.Style = "Заголовок 2"
    if ($Mode -eq "draft") {
        $selection.TypeText("Ожидаемые расчётные показатели (маски)")
    } else {
        $selection.TypeText("Результаты расчётов")
    }
    $selection.TypeParagraph()
    $selection.Style = "Обычный"

    $table2 = $doc.Tables.Add($selection.Range, 7, 2)
    $table2.Borders.Enable = 1
    for ($row = 1; $row -le 7; $row++) {
        for ($col = 1; $col -le 2; $col++) {
            $table2.Cell($row, $col).Range.ParagraphFormat.Alignment = 1
        }
    }
    $table2.Cell(1,1).Range.Text = "Показатель"; $table2.Cell(1,1).Range.Font.Bold = $true
    $table2.Cell(1,2).Range.Text = "Значение"; $table2.Cell(1,2).Range.Font.Bold = $true

    if ($Mode -eq "draft") {
        $table2.Cell(2,1).Range.Text = "Количество измерений"; $table2.Cell(2,2).Range.Text = "[n]"
        $table2.Cell(3,1).Range.Text = "Среднее арифметическое (площадь)"; $table2.Cell(3,2).Range.Text = "[x.xxx]"
        $table2.Cell(4,1).Range.Text = "Стандартное отклонение (СКО)"; $table2.Cell(4,2).Range.Text = "[x.xxx]"
        $table2.Cell(5,1).Range.Text = "Относительное СКО, %"; $table2.Cell(5,2).Range.Text = "[x.xxx]"
        $table2.Cell(6,1).Range.Text = "Минимум"; $table2.Cell(6,2).Range.Text = "[x.xxx]"
        $table2.Cell(7,1).Range.Text = "Максимум"; $table2.Cell(7,2).Range.Text = "[x.xxx]"
    } else {
        $table2.Cell(2,1).Range.Text = "Количество измерений"; $table2.Cell(2,2).Range.Text = $count.ToString()
        $table2.Cell(3,1).Range.Text = "Среднее арифметическое (площадь)"; $table2.Cell(3,2).Range.Text = [math]::Round($mean,4).ToString()
        $table2.Cell(4,1).Range.Text = "Стандартное отклонение (СКО)"; $table2.Cell(4,2).Range.Text = [math]::Round($stddev,4).ToString()
        $table2.Cell(5,1).Range.Text = "Относительное СКО, %"; $table2.Cell(5,2).Range.Text = [math]::Round($rsd,4).ToString()
        $table2.Cell(6,1).Range.Text = "Минимум"; $table2.Cell(6,2).Range.Text = [math]::Round($min,4).ToString()
        $table2.Cell(7,1).Range.Text = "Максимум"; $table2.Cell(7,2).Range.Text = [math]::Round($max,4).ToString()
    }
    $selection.EndKey(6) | Out-Null
    #$selection.TypeParagraph()

    # Вставка графика (из буфера) с масштабированием
    $selection.Style = "Заголовок 2"
    if ($Mode -eq "draft") {
        $selection.TypeText("Графическое представление (шаблон)")
    } else {
        $selection.TypeText("Графическое представление данных")
    }
    $selection.TypeParagraph()
    $selection.Style = "Обычный"

    # Отключаем отступ первой строки и центрируем для графика
    $selection.ParagraphFormat.FirstLineIndent = 0
    $selection.ParagraphFormat.Alignment = 1   # 1 = центрирование

    # Вставляем график
    $selection.PasteSpecial($null, $false, 0, $false, 0)
    Start-Sleep -Milliseconds 300

    # Масштабируем вставленное изображение (InlineShape)
    $shape = $doc.InlineShapes.Item($doc.InlineShapes.Count)
    $pageSetup = $doc.PageSetup
    $usableWidth = $pageSetup.PageWidth - $pageSetup.LeftMargin - $pageSetup.RightMargin
    $targetWidth = $usableWidth * ($PercentWidth / 100)
    $ratio = ($targetWidth / $shape.Width) * 100
    $shape.LockAspectRatio = -1
    $shape.ScaleWidth = $ratio
    $shape.ScaleHeight = $ratio

    $selection.TypeParagraph()

    # Вердикт
    $selection.Style = "Заголовок 2"
    if ($Mode -eq "draft") {
        $selection.TypeText("Предварительное заключение")
    } else {
        $selection.TypeText("Вердикт о прецизионности")
    }
    $selection.TypeParagraph()
    $selection.Style = "Обычный"

    if ($Mode -eq "draft") {
        $selection.TypeText("Прецизионность ")
        $selection.Font.Bold = $true
        $selection.TypeText("соответствует / не соответствует")
        $selection.Font.Bold = $false
        $selection.TypeText(" критерию.")
    } else {
        if ($rsd -le 2.0) {
            $selection.TypeText("Прецизионность ")
            $selection.Font.Bold = $true
            $selection.TypeText("соответствует")
            $selection.Font.Bold = $false
            $selection.TypeText(" критерию (ОСКО ≤ 2%)")
        } else {
            $selection.TypeText("Прецизионность ")
            $selection.Font.Bold = $true
            $selection.TypeText("НЕ соответствует")
            $selection.Font.Bold = $false
            $selection.TypeText(" критерию (ОСКО > 2%)")
        }
    }

    # Сохранение Word
    if (Test-Path $OutputDocx) { Remove-Item $OutputDocx -Force }
    $doc.SaveAs([string]$OutputDocx)
    $doc.Close()
    $word.Quit()
    $wb.Close($false)
    $excel.Quit()

    Write-Host "DOCX сохранён: $OutputDocx" -ForegroundColor Green
    Start-Process $OutputDocx
}
catch {
    Write-Host "ОШИБКА: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    # Освобождение ресурсов
    if ($excel) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
    if ($word)  { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)  | Out-Null }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
