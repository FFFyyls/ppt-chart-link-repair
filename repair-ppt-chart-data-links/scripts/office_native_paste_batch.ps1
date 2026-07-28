# Copyright 2026 XLoffice-Fyl
# SPDX-License-Identifier: Apache-2.0

param(
    [Parameter(Mandatory=$true)][string]$ManifestPath,
    [Parameter(Mandatory=$true)][string]$SupportMapPath,
    [Parameter(Mandatory=$true)][string]$CandidatePptx,
    [Parameter(Mandatory=$true)][string]$PasteDir,
    [Parameter(Mandatory=$true)][string]$BatchResultPath
)

$ErrorActionPreference = 'Stop'
$startedAt = Get-Date
$manifest = [IO.File]::ReadAllText([IO.Path]::GetFullPath($ManifestPath), [Text.Encoding]::UTF8) | ConvertFrom-Json
$supportMap = [IO.File]::ReadAllText([IO.Path]::GetFullPath($SupportMapPath), [Text.Encoding]::UTF8) | ConvertFrom-Json
$candidate = [IO.Path]::GetFullPath($CandidatePptx)
$PasteDir = [IO.Path]::GetFullPath($PasteDir)
$BatchResultPath = [IO.Path]::GetFullPath($BatchResultPath)
New-Item -ItemType Directory -Force -Path $PasteDir | Out-Null
$progressPath = Join-Path $PasteDir 'batch-progress.log'
[IO.File]::WriteAllText($progressPath, '', [Text.UTF8Encoding]::new($false))
$maxPasteAttempts = 2

function Get-LivePowerPointPids {
    return @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object {
        try { $_.Threads.Count -gt 0 } catch { $false }
    } | ForEach-Object { $_.Id })
}

if ((Get-LivePowerPointPids).Count -gt 0) {
    [Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
    Write-Output '{"passed":false,"requires_user_action":true,"message":"PowerPoint正在运行，请保存并关闭所有PowerPoint窗口后继续"}'
    exit 20
}

$baselinePowerPointPids = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$baselineExcelPids = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$excel = $null; $powerpoint = $null; $presentation = $null; $window = $null
$excelPid = $null; $powerpointPid = $null
$books = @{}
$results = @()
$batchSaved = $false

Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class NativePasteBatchWindowPid {
  [DllImport("user32.dll")]
  public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
"@

function Release-Com($value) {
    if ($null -ne $value) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($value) } catch {}
    }
}

function Get-OfficePid($application) {
    try {
        [uint32]$processId = 0
        [void][NativePasteBatchWindowPid]::GetWindowThreadProcessId([IntPtr][int64]$application.Hwnd, [ref]$processId)
        return [int]$processId
    } catch { return $null }
}

function Stop-OwnedProcess([Nullable[int]]$processId, [int[]]$baselinePids) {
    if ($null -eq $processId -or $processId -in $baselinePids) { return }
    if ($null -ne (Get-Process -Id $processId -ErrorAction SilentlyContinue)) {
        Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
    }
}

function Stop-OwnedOfficeDescendants([int]$parentProcessId, [int[]]$baselinePids) {
    try {
        $children = @(Get-CimInstance Win32_Process -Filter "ParentProcessId=$parentProcessId" -ErrorAction Stop)
        foreach ($child in $children) {
            if ([string]$child.Name -notin @('EXCEL.EXE', 'POWERPNT.EXE')) { continue }
            $childPid = [int]$child.ProcessId
            if ($childPid -in $baselinePids) { continue }
            Stop-Process -Id $childPid -Force -ErrorAction SilentlyContinue
        }
    } catch {}
}

function Get-ShapeByName($slide, [string]$name) {
    $fallbackName = $name
    if ($name -match '(\d+)$') { $fallbackName = 'Chart ' + $Matches[1] }
    for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
        $shape = $slide.Shapes.Item($i)
        if ([string]$shape.Name -eq $name -or [string]$shape.Name -eq $fallbackName) { return $shape }
        Release-Com $shape
    }
    throw "Shape not found: $name"
}

function Get-NewShape($slide, $beforeIds) {
    for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
        $shape = $slide.Shapes.Item($i)
        if (-not $beforeIds.Contains([int]$shape.Id)) { return $shape }
        Release-Com $shape
    }
    return $null
}

function Get-RuntimeSeriesNames($chart) {
    $collection = $null
    $names = @()
    try {
        $collection = $chart.SeriesCollection()
        for ($index = 1; $index -le [int]$collection.Count; $index++) {
            $series = $null
            try {
                $series = $collection.Item($index)
                $names += [string]$series.Name
            } finally { Release-Com $series }
        }
    } finally { Release-Com $collection }
    return @($names)
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excelPid = Get-OfficePid $excel
    if ($null -eq $excelPid -or $excelPid -in $baselineExcelPids) {
        throw 'Cannot establish an isolated Excel process for native batch paste.'
    }
    $excel.Visible = $true
    $excel.DisplayAlerts = $false

    $powerpoint = New-Object -ComObject PowerPoint.Application
    $powerpointPid = Get-OfficePid $powerpoint
    if ($null -eq $powerpointPid -or $powerpointPid -in $baselinePowerPointPids) {
        throw 'Cannot establish an isolated PowerPoint process for native batch paste.'
    }
    $powerpoint.Visible = -1
    $powerpoint.DisplayAlerts = 1
    $presentation = $powerpoint.Presentations.Open($candidate, $false, $false, $true)
    if ($powerpoint.Windows.Count -eq 0) { [void]$presentation.NewWindow() }
    $window = $powerpoint.Windows.Item(1)

    foreach ($chart in @($manifest.charts)) {
        $chartId = [string]$chart.chart_id
        Add-Content -LiteralPath $progressPath -Encoding UTF8 -Value ((Get-Date).ToString('o') + "`tchart:start:" + $chartId)
        $support = @($supportMap | Where-Object { [string]$_.chart_id -eq $chartId })[0]
        if ($null -eq $support) { throw "Support map not found: $chartId" }
        $workbookPath = [IO.Path]::GetFullPath([string]$support.workbook_path)
        if (-not $books.ContainsKey($workbookPath)) {
            $books[$workbookPath] = $excel.Workbooks.Open($workbookPath, 0, $false)
        }
        $workbook = $books[$workbookPath]
        $sheet = $null; $chartObject = $null; $slide = $null
        $oldShape = $null; $newShape = $null
        $linkedChart = $null; $linkedChartData = $null; $linkedLinkFormat = $null
        try {
            $sheet = $workbook.Worksheets.Item([string]$support.support_sheet)
            $chartObject = $sheet.ChartObjects().Item([string]$support.chart_object)
            $slide = $presentation.Slides.Item([int]$chart.actual_slide)
            $oldShape = Get-ShapeByName $slide ([string]$chart.shape_name)
            $geometry = [ordered]@{
                left = [double]$oldShape.Left
                top = [double]$oldShape.Top
                width = [double]$oldShape.Width
                height = [double]$oldShape.Height
                z = [int]$oldShape.ZOrderPosition
            }
            $beforeIds = [Collections.Generic.HashSet[int]]::new()
            for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
                $shape = $slide.Shapes.Item($i)
                [void]$beforeIds.Add([int]$shape.Id)
                Release-Com $shape
            }

            for ($pasteAttempt = 1; $pasteAttempt -le $maxPasteAttempts -and $null -eq $newShape; $pasteAttempt++) {
                if ($pasteAttempt -gt 1) {
                    $newShape = Get-NewShape $slide $beforeIds
                    if ($null -ne $newShape) { break }
                }
                Add-Content -LiteralPath $progressPath -Encoding UTF8 -Value ((Get-Date).ToString('o') + "`tpaste:attempt:" + $chartId + ":" + $pasteAttempt)
                $sheet.Activate()
                $chartObject.Activate()
                $excel.CommandBars.ExecuteMso('Copy')
                Start-Sleep -Milliseconds 250
                $window.Activate()
                $window.View.GotoSlide([int]$chart.actual_slide)
                $powerpoint.CommandBars.ExecuteMso('PasteSourceFormatting')

                $deadline = (Get-Date).AddSeconds(20)
                do {
                    Start-Sleep -Milliseconds 100
                    $newShape = Get-NewShape $slide $beforeIds
                } while ($null -eq $newShape -and (Get-Date) -lt $deadline)
            }
            if ($null -eq $newShape) { throw "Pasted native chart did not materialize: $chartId" }
            if ($newShape.HasChart -ne -1) { throw "Pasted shape is not a native chart: $chartId" }
            $linkedChart = $newShape.Chart
            $linkedChartData = $linkedChart.ChartData
            $linkedLinkFormat = $newShape.LinkFormat
            if (-not [bool]$linkedChartData.IsLinked) { throw "Pasted chart is not externally linked: $chartId" }

            $newShape.Left = $geometry.left
            $newShape.Top = $geometry.top
            $newShape.Width = $geometry.width
            $newShape.Height = $geometry.height
            $oldShape.Delete()
            $newShapeName = "__NATIVE_LINK_" + ($chartId -replace '[^A-Za-z0-9]', '_')
            $newShape.Name = $newShapeName
            $runtimeNames = @(Get-RuntimeSeriesNames $linkedChart)
            $results += [ordered]@{
                chart_id = $chartId
                actual_slide = [int]$chart.actual_slide
                original_shape_name = [string]$chart.shape_name
                linked_shape_name = $newShapeName
                shape_id_before_close = [int]$newShape.Id
                linked = [bool]$linkedChartData.IsLinked
                has_chart = [int]$newShape.HasChart
                chart_type_before_close = [int]$linkedChart.ChartType
                link_source_before_close = [string]$linkedLinkFormat.SourceFullName
                runtime_series_names = @($runtimeNames)
                geometry = $geometry
                saved = $false
                passed = $true
            }
            Add-Content -LiteralPath $progressPath -Encoding UTF8 -Value ((Get-Date).ToString('o') + "`tchart:passed:" + $chartId)
        } finally {
            Release-Com $linkedLinkFormat; Release-Com $linkedChartData; Release-Com $linkedChart
            Release-Com $newShape; Release-Com $oldShape; Release-Com $slide
            Release-Com $chartObject; Release-Com $sheet
        }
    }

    $presentation.Save()
    $batchSaved = $true
    foreach ($result in $results) {
        $result.saved = $true
        $resultPath = Join-Path $PasteDir (([string]$result.chart_id) + '.json')
        [IO.File]::WriteAllText($resultPath, ($result | ConvertTo-Json -Depth 8), [Text.UTF8Encoding]::new($false))
    }
    $batch = [ordered]@{
        passed = $true
        chart_count = @($results).Count
        powerpoint_pid = $powerpointPid
        excel_pid = $excelPid
        powerpoint_launch_count = 1
        excel_launch_count = 1
        elapsed_seconds = [Math]::Round(((Get-Date) - $startedAt).TotalSeconds, 3)
    }
    [IO.File]::WriteAllText($BatchResultPath, ($batch | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false))
    $batch | ConvertTo-Json -Depth 6
}
finally {
    Release-Com $window
    if ($null -ne $presentation) {
        try { if (-not $batchSaved) { $presentation.Saved = -1 }; $presentation.Close() } catch {}
        Release-Com $presentation
    }
    if ($null -ne $powerpoint -and $null -ne $powerpointPid -and $powerpointPid -notin $baselinePowerPointPids) {
        try { $powerpoint.Quit() } catch {}
    }
    Release-Com $powerpoint
    foreach ($book in $books.Values) { try { $book.Close($false) } catch {}; Release-Com $book }
    if ($null -ne $excel -and $null -ne $excelPid -and $excelPid -notin $baselineExcelPids) {
        try { $excel.Quit() } catch {}
    }
    Release-Com $excel
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 350
    Stop-OwnedProcess $powerpointPid $baselinePowerPointPids
    Stop-OwnedProcess $excelPid $baselineExcelPids
    Stop-OwnedOfficeDescendants $PID @($baselinePowerPointPids + $baselineExcelPids)
}
