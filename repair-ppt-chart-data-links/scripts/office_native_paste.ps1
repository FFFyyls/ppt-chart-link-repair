# Copyright 2026 XLoffice-Fyl
# SPDX-License-Identifier: Apache-2.0

param(
    [Parameter(Mandatory=$true)][string]$ManifestPath,
    [Parameter(Mandatory=$true)][string]$SupportMapPath,
    [Parameter(Mandatory=$true)][string]$CandidatePptx,
    [Parameter(Mandatory=$true)][string]$ChartId,
    [Parameter(Mandatory=$true)][string]$ResultPath
)

$ErrorActionPreference = 'Stop'
$manifest = [IO.File]::ReadAllText([IO.Path]::GetFullPath($ManifestPath), [Text.Encoding]::UTF8) | ConvertFrom-Json
$supportMap = [IO.File]::ReadAllText([IO.Path]::GetFullPath($SupportMapPath), [Text.Encoding]::UTF8) | ConvertFrom-Json
$chart = @($manifest.charts | Where-Object { [string]$_.chart_id -eq $ChartId })[0]
$support = @($supportMap | Where-Object { [string]$_.chart_id -eq $ChartId })[0]
if ($null -eq $chart -or $null -eq $support) { throw "Chart or support map not found: $ChartId" }
$baselinePowerPointPids = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$baselineExcelPids = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })

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

$candidate = [IO.Path]::GetFullPath($CandidatePptx)
$excel = $null; $workbook = $null; $sheet = $null; $chartObject = $null
$powerpoint = $null; $presentation = $null
$slide = $null; $oldShape = $null; $newShape = $null; $freshShape = $null; $window = $null
$excelPid = $null; $powerpointPid = $null
$newShapeName = "__NATIVE_LINK_" + ($ChartId -replace '[^A-Za-z0-9]', '_')

Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class NativePasteWindowPid {
  [DllImport("user32.dll")]
  public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
"@

function Release-Com($value) {
    if ($null -ne $value) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($value) } catch {}
    }
}

function Get-ExcelPid($application) {
    try {
        [uint32]$processId = 0
        [void][NativePasteWindowPid]::GetWindowThreadProcessId([IntPtr][int64]$application.Hwnd, [ref]$processId)
        return [int]$processId
    } catch { return $null }
}

function Get-PowerPointPid($application) {
    try {
        [uint32]$processId = 0
        [void][NativePasteWindowPid]::GetWindowThreadProcessId([IntPtr][int64]$application.Hwnd, [ref]$processId)
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

try {
    $excel = New-Object -ComObject Excel.Application
    $excelPid = Get-ExcelPid $excel
    if ($null -eq $excelPid -or $excelPid -in $baselineExcelPids) {
        throw 'Cannot establish an isolated Excel process for native paste.'
    }
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open([string]$support.workbook_path, 0, $false)
    $sheet = $workbook.Worksheets.Item([string]$support.support_sheet)
    $chartObject = $sheet.ChartObjects().Item([string]$support.chart_object)

    $powerpoint = New-Object -ComObject PowerPoint.Application
    $powerpointPid = Get-PowerPointPid $powerpoint
    if ($null -eq $powerpointPid -or $powerpointPid -in $baselinePowerPointPids) {
        throw 'Cannot establish an isolated PowerPoint process for native paste.'
    }
    $powerpoint.Visible = -1
    $powerpoint.DisplayAlerts = 1
    $presentation = $powerpoint.Presentations.Open($candidate, $false, $false, $true)
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

    $sheet.Activate()
    $chartObject.Activate()
    $excel.CommandBars.ExecuteMso('Copy')
    Start-Sleep -Milliseconds 800

    if ($powerpoint.Windows.Count -eq 0) { [void]$presentation.NewWindow() }
    $window = $powerpoint.Windows.Item(1)
    $window.Activate()
    $window.View.GotoSlide([int]$chart.actual_slide)
    $powerpoint.CommandBars.ExecuteMso('PasteSourceFormatting')

    $deadline = (Get-Date).AddSeconds(20)
    $newShape = $null
    do {
        Start-Sleep -Milliseconds 250
        $newShape = Get-NewShape $slide $beforeIds
    } while ($null -eq $newShape -and (Get-Date) -lt $deadline)
    if ($null -eq $newShape) { throw "Pasted native chart did not materialize: $ChartId" }
    if ($newShape.HasChart -ne -1) { throw "Pasted shape is not a native chart: $ChartId" }
    if (-not [bool]$newShape.Chart.ChartData.IsLinked) { throw "Pasted chart is not externally linked: $ChartId" }

    $newShape.Left = $geometry.left
    $newShape.Top = $geometry.top
    $newShape.Width = $geometry.width
    $newShape.Height = $geometry.height
    $oldShape.Delete()
    $newShape.Name = $newShapeName
    $linkTargetBeforeClose = [string]$newShape.LinkFormat.SourceFullName
    $shapeIdBeforeClose = [int]$newShape.Id
    $presentation.Save()
    $result = [ordered]@{
        chart_id = $ChartId
        actual_slide = [int]$chart.actual_slide
        original_shape_name = [string]$chart.shape_name
        linked_shape_name = $newShapeName
        shape_id_before_close = $shapeIdBeforeClose
        linked = [bool]$newShape.Chart.ChartData.IsLinked
        has_chart = [int]$newShape.HasChart
        chart_type_before_close = [int]$newShape.Chart.ChartType
        link_source_before_close = $linkTargetBeforeClose
        geometry = $geometry
        saved = $true
        passed = $true
    }
    [IO.File]::WriteAllText([IO.Path]::GetFullPath($ResultPath), ($result | ConvertTo-Json -Depth 8), [Text.UTF8Encoding]::new($false))
}
finally {
    Release-Com $freshShape
    Release-Com $newShape
    Release-Com $oldShape
    Release-Com $slide
    Release-Com $window
    if ($null -ne $presentation) { try { $presentation.Saved = -1; $presentation.Close() } catch {}; Release-Com $presentation }
    if ($null -ne $powerpoint -and $null -ne $powerpointPid -and $powerpointPid -notin $baselinePowerPointPids) { try { $powerpoint.Quit() } catch {} }
    Release-Com $powerpoint
    if ($null -ne $chartObject) { Release-Com $chartObject }
    if ($null -ne $sheet) { Release-Com $sheet }
    if ($null -ne $workbook) { try { $workbook.Close($true) } catch {}; Release-Com $workbook }
    if ($null -ne $excel -and $null -ne $excelPid -and $excelPid -notin $baselineExcelPids) { try { $excel.Quit() } catch {} }
    Release-Com $excel
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 350
    Stop-OwnedProcess $powerpointPid $baselinePowerPointPids
    Stop-OwnedProcess $excelPid $baselineExcelPids
}
