# Copyright 2026 XLoffice-Fyl
# SPDX-License-Identifier: Apache-2.0

param(
    [Parameter(Mandatory=$true)][string]$ManifestPath,
    [Parameter(Mandatory=$true)][string]$ResultPath
)

$ErrorActionPreference = 'Stop'
$manifest = [IO.File]::ReadAllText([IO.Path]::GetFullPath($ManifestPath), [Text.Encoding]::UTF8) | ConvertFrom-Json
$excel = $null
$excelPid = $null
$baselineExcelPids = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$books = @{}
$results = @()
$runToken = Get-Date -Format 'HHmmss'

Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class NativeCarrierWindowPid {
  [DllImport("user32.dll")]
  public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
"@

function Get-ExcelPid($application) {
    try {
        [uint32]$processId = 0
        [void][NativeCarrierWindowPid]::GetWindowThreadProcessId([IntPtr][int64]$application.Hwnd, [ref]$processId)
        return [int]$processId
    } catch { return $null }
}

function Stop-OwnedProcess([Nullable[int]]$processId, [int[]]$baselinePids) {
    if ($null -eq $processId -or $processId -in $baselinePids) { return }
    $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
    if ($null -ne $process) {
        Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
    }
}

function Chart-Type-Value($series) {
    $plot = [string]$series.plot_type
    $grouping = [string]$series.grouping
    $barDir = [string]$series.bar_dir
    if ($plot -eq 'barChart' -and [string]::IsNullOrWhiteSpace($barDir)) { $barDir = 'col' }
    if ($plot -eq 'barChart' -and [string]::IsNullOrWhiteSpace($grouping)) { $grouping = 'clustered' }
    if (($plot -eq 'lineChart' -or $plot -eq 'areaChart') -and [string]::IsNullOrWhiteSpace($grouping)) { $grouping = 'standard' }
    if ($plot -eq 'barChart' -and $barDir -eq 'col') {
        if ($grouping -eq 'stacked') { return 52 }
        if ($grouping -eq 'percentStacked') { return 53 }
        return 51
    }
    if ($plot -eq 'barChart' -and $barDir -eq 'bar') {
        if ($grouping -eq 'stacked') { return 58 }
        if ($grouping -eq 'percentStacked') { return 59 }
        return 57
    }
    if ($plot -eq 'lineChart') {
        if ($grouping -eq 'stacked') { return 63 }
        if ($grouping -eq 'percentStacked') { return 64 }
        return 4
    }
    if ($plot -eq 'areaChart') {
        if ($grouping -eq 'stacked') { return 76 }
        if ($grouping -eq 'percentStacked') { return 77 }
        return 1
    }
    if ($plot -eq 'scatterChart') { return -4169 }
    if ($plot -eq 'pieChart') { return 5 }
    if ($plot -eq 'doughnutChart') { return -4120 }
    if ($plot -eq 'radarChart') { return -4151 }
    if ($plot -eq 'bubbleChart') { return 15 }
    if ($plot -eq 'surfaceChart') { return 83 }
    if ($plot -eq 'stockChart') { return 88 }
    if ($plot -eq 'ofPieChart') { return 68 }
    throw "Unsupported plot type for native carrier: $plot/$grouping/$barDir"
}

function Get-Range($workbook, $parsed) {
    $sheet = $workbook.Worksheets.Item([string]$parsed.sheet)
    return $sheet.Range([string]$parsed.address)
}

function Get-Formula-Address($parsed) {
    $address = [string]$parsed.address
    $address = [regex]::Replace(
        $address,
        '([A-Z]+)(\d+)',
        { param($match) '$' + $match.Groups[1].Value + '$' + $match.Groups[2].Value }
    )
    $escapedSheet = ([string]$parsed.sheet) -replace "'", "''"
    return "='" + $escapedSheet + "'!" + $address
}

function Release-Com($value) {
    if ($null -ne $value) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($value) } catch {}
    }
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excelPid = Get-ExcelPid $excel
    if ($null -eq $excelPid -or $excelPid -in $baselineExcelPids) {
        throw 'Cannot establish an isolated Excel process for native carrier generation.'
    }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    foreach ($chart in @($manifest.charts)) {
        $chartId = [string]$chart.chart_id
        Write-Output ("chart:start:{0}" -f $chartId)
        $workbookPath = [IO.Path]::GetFullPath([string]$chart.workbook_path)
        if (-not $books.ContainsKey($workbookPath)) {
            $books[$workbookPath] = $excel.Workbooks.Open($workbookPath, 0, $false)
        }
        $workbook = $books[$workbookPath]
        $sheetName = ("L_" + ($chartId -replace '[^A-Za-z0-9]', '_') + "_" + $runToken)
        if ($sheetName.Length -gt 31) { $sheetName = $sheetName.Substring(0, 31) }
        $sheet = $workbook.Worksheets.Add()
        $sheet.Name = $sheetName
        $chartObject = $sheet.ChartObjects().Add(20, 20, 640, 360)
        $chartObject.Name = "LinkCarrier"
        $firstType = Chart-Type-Value $chart.series[0]
        $chartObject.Chart.ChartType = $firstType
        while ($chartObject.Chart.SeriesCollection().Count -gt 0) {
            $chartObject.Chart.SeriesCollection(1).Delete()
        }

        foreach ($recipeSeries in @($chart.series | Sort-Object {[int]$_.order})) {
            Write-Output ("series:start:{0}:{1}:{2}" -f $chartId, $recipeSeries.order, $recipeSeries.plot_type)
            $series = $chartObject.Chart.SeriesCollection().NewSeries()
            $refs = $recipeSeries.refs
            if ($null -ne $refs.tx) {
                if ([string]$refs.tx.parsed.kind -eq 'defined-name') {
                    Write-Output ("series:name-defined:{0}:{1}" -f $chartId, $refs.tx.parsed.name)
                    $definedName = $workbook.Names.Item([string]$refs.tx.parsed.name)
                    $definedNameRange = $definedName.RefersToRange
                    $definedSheet = ([string]$definedNameRange.Worksheet.Name) -replace "'", "''"
                    $definedAddress = [string]$definedNameRange.Address($true, $true, 1, $false)
                    $series.Name = "='" + $definedSheet + "'!" + $definedAddress
                    Release-Com $definedNameRange
                    Release-Com $definedName
                } else {
                    Write-Output ("series:name-range:{0}:{1}!{2}" -f $chartId, $refs.tx.parsed.sheet, $refs.tx.parsed.address)
                    $series.Name = Get-Formula-Address $refs.tx.parsed
                }
            }
            $categoryRef = if ($null -ne $refs.xVal) { $refs.xVal } else { $refs.cat }
            if ($null -ne $categoryRef -and [string]$categoryRef.parsed.kind -eq 'range') {
                Write-Output ("series:xvalues:{0}:{1}!{2}" -f $chartId, $categoryRef.parsed.sheet, $categoryRef.parsed.address)
                $series.XValues = Get-Formula-Address $categoryRef.parsed
            }
            $valueRef = if ($null -ne $refs.val) { $refs.val } else { $refs.yVal }
            if ($null -ne $valueRef -and [string]$valueRef.parsed.kind -eq 'range') {
                Write-Output ("series:values:{0}:{1}!{2}" -f $chartId, $valueRef.parsed.sheet, $valueRef.parsed.address)
                $series.Values = Get-Formula-Address $valueRef.parsed
            }
            if ($null -ne $refs.bubbleSize -and [string]$refs.bubbleSize.parsed.kind -eq 'range') {
                $series.BubbleSizes = Get-Formula-Address $refs.bubbleSize.parsed
            }
            $desiredType = Chart-Type-Value $recipeSeries
            $currentType = $null
            try { $currentType = [int]$series.ChartType } catch {}
            Write-Output ("series:type-request:{0}:current={1}:desired={2}" -f $chartId, $currentType, $desiredType)
            if ($currentType -ne $desiredType) {
                $series.ChartType = $desiredType
            }
            Write-Output ("series:type:{0}:{1}" -f $chartId, $desiredType)
            $axisGroup = 1
            if ($recipeSeries.PSObject.Properties.Name -contains 'axis_group') {
                $axisGroup = [int]$recipeSeries.axis_group
            }
            try { $series.AxisGroup = $axisGroup } catch {}
            Release-Com $series
        }

        $sheet.Activate()
        $chartObject.Activate()
        $results += [ordered]@{
            chart_id = $chartId
            workbook_path = $workbookPath
            support_sheet = $sheetName
            chart_object = 'LinkCarrier'
            series_count = [int]$chartObject.Chart.SeriesCollection().Count
            chart_type = [int]$chartObject.Chart.ChartType
        }
        Release-Com $chartObject
        Release-Com $sheet
        Write-Output ("chart:done:{0}" -f $chartId)
    }

    foreach ($book in $books.Values) { $book.Save() }
    [IO.File]::WriteAllText(
        [IO.Path]::GetFullPath($ResultPath),
        ($results | ConvertTo-Json -Depth 8),
        [Text.UTF8Encoding]::new($false)
    )
}
finally {
    foreach ($book in $books.Values) { try { $book.Close($true) } catch {}; Release-Com $book }
    if ($null -ne $excel -and $null -ne $excelPid -and $excelPid -notin $baselineExcelPids) { try { $excel.Quit() } catch {} }
    Release-Com $excel
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 350
    Stop-OwnedProcess $excelPid $baselineExcelPids
}
