# Copyright 2026 XLoffice-Fyl
# SPDX-License-Identifier: Apache-2.0

param(
  [Parameter(Mandatory=$true)][ValidateSet('Primary','EditData','WorkbookPath','MutationUpdate','ReadValue')][string]$Mode,
  [Parameter(Mandatory=$true)][string]$ManifestPath,
  [Parameter(Mandatory=$true)][string]$ChartId,
  [Parameter(Mandatory=$true)][string]$Pptx,
  [Parameter(Mandatory=$true)][string]$ExpectedWorkbook,
  [string]$ExpectedSeriesName = '',
  [int]$SeriesIndex = 0,
  [int]$PointIndex = 0,
  [double]$ExpectedValue = [double]::NaN,
  [switch]$SaveAfterUpdate
)

$ErrorActionPreference = 'Stop'
$manifest = [IO.File]::ReadAllText([IO.Path]::GetFullPath($ManifestPath), [Text.Encoding]::UTF8) | ConvertFrom-Json
$chartInfo = @($manifest.charts | Where-Object { [string]$_.chart_id -eq $ChartId })[0]
if ($null -eq $chartInfo) { throw "Chart not found: $ChartId" }
$Pptx = [IO.Path]::GetFullPath($Pptx)
$ExpectedWorkbook = [IO.Path]::GetFullPath($ExpectedWorkbook)
$powerpoint = $null; $presentation = $null; $slide = $null; $shape = $null; $chartObject = $null
$chartData = $null; $excelWorkbook = $null; $activatedWorkbook = $null; $excelApplication = $null; $editDataApplication = $null; $excelPid = $null
$powerpointPid = $null
$updateExcelPid = $null
$excelOwnedByWorker = $false
$probeSheet = $null; $probeCell = $null

function Get-LivePowerPointPids {
  @(
    Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object {
      try { $_.Threads.Count -gt 0 } catch { $true }
    } | ForEach-Object { $_.Id }
  )
}

$pidsBeforePowerPoint = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$pidsBeforeExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
if (@(Get-LivePowerPointPids).Count -gt 0) {
  [ordered]@{
    passed = $false
    requires_user_action = $true
    message = 'PowerPoint正在运行，请保存并关闭所有PowerPoint窗口后继续'
  } | ConvertTo-Json -Depth 4
  exit 20
}
$progressPath = Join-Path ([IO.Path]::GetDirectoryName($Pptx)) ($ChartId + '-' + $Mode + '-progress.log')
[IO.File]::WriteAllText($progressPath, "start`r`n", [Text.UTF8Encoding]::new($false))

function Mark-Progress([string]$stage) {
  Add-Content -LiteralPath $progressPath -Encoding UTF8 -Value ((Get-Date).ToString('o') + "`t" + $stage)
}

Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class StageWorkerWindowPid {
  [DllImport("user32.dll")]
  public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
"@

function Release-Com($object) {
  if ($null -ne $object) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($object) } catch {} }
}

function Get-OfficePid($application) {
  try {
    [uint32]$processId = 0
    [void][StageWorkerWindowPid]::GetWindowThreadProcessId([IntPtr][int64]$application.Hwnd, [ref]$processId)
    return [int]$processId
  } catch { return $null }
}

function Resolve-OwnedUpdateExcelPid([int[]]$pidsBeforeUpdate) {
  $deadline = (Get-Date).AddSeconds(3)
  do {
    $candidates = @(
      Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object {
        if ($_.Id -in $pidsBeforeUpdate) { return $false }
        try { return $_.Threads.Count -gt 0 } catch { return $false }
      } | ForEach-Object { $_.Id }
    )
    if ($candidates.Count -eq 1) { return [int]$candidates[0] }
    if ($candidates.Count -gt 1) {
      throw 'Ambiguous Excel processes appeared during LinkFormat.Update; no process will be terminated.'
    }
    Start-Sleep -Milliseconds 100
  } while ((Get-Date) -lt $deadline)
  return $null
}

function Resolve-SeriesOrdinal($chart, [string]$expectedName, [int]$fallbackZeroBased) {
  if ([string]::IsNullOrWhiteSpace($expectedName)) {
    throw 'A unique runtime series name is required; numeric ordinal fallback is forbidden.'
  }
  $collection = $null
  $matches = @()
  try {
    $collection = $chart.SeriesCollection()
    for ($index = 1; $index -le [int]$collection.Count; $index++) {
      $candidate = $null
      try {
        $candidate = $collection.Item($index)
        if ([string]::Equals([string]$candidate.Name, $expectedName, [StringComparison]::Ordinal)) {
          $matches += $index
        }
      } finally { Release-Com $candidate }
    }
  } finally { Release-Com $collection }
  if ($matches.Count -ne 1) {
    throw "Expected exactly one COM series named '$expectedName'; found $($matches.Count)."
  }
  return [int]$matches[0]
}

function Get-PointValue($chart, [int]$seriesOneBased, [int]$pointZeroBased) {
  $series = $null
  try {
    $series = $chart.SeriesCollection($seriesOneBased)
    $values = @($series.Values)
    if ($pointZeroBased -lt 0 -or $pointZeroBased -ge $values.Count) {
      throw "Point index $pointZeroBased outside 0..$($values.Count - 1)"
    }
    return [double]$values[$pointZeroBased]
  } finally { Release-Com $series }
}

function Get-TargetChartShape($slide, [string]$expectedName) {
  try { return $slide.Shapes.Item($expectedName) } catch {}
  $suffix = $null
  if ($expectedName -match '(\d+)\s*$') { $suffix = $Matches[1] }
  $matches = @()
  for ($index = 1; $index -le $slide.Shapes.Count; $index++) {
    $candidate = $null
    try {
      $candidate = $slide.Shapes.Item($index)
      if ($candidate.HasChart -eq -1) {
        $nameMatches = ($null -ne $suffix -and [string]$candidate.Name -match ("{0}\s*$" -f [regex]::Escape($suffix)))
        if ($nameMatches) { $matches += $index }
      }
    } finally { Release-Com $candidate }
  }
  if ($matches.Count -eq 1) { return $slide.Shapes.Item([int]$matches[0]) }
  throw "Cannot resolve native chart shape '$expectedName' on slide $($slide.SlideIndex)."
}

try {
  $powerpoint = New-Object -ComObject PowerPoint.Application
  $powerpointPid = Get-OfficePid $powerpoint
  if ($null -eq $powerpointPid) { throw 'Cannot determine the PowerPoint process created by the worker.' }
  if ($powerpointPid -in $pidsBeforePowerPoint) {
    throw 'PowerPoint automation attached to a pre-existing process; the worker will not use or close it.'
  }
  Mark-Progress ("powerpoint-pid=" + $powerpointPid)
  $powerpoint.Visible = -1
  $powerpoint.DisplayAlerts = 2
  $willSave = (($Mode -eq 'MutationUpdate' -or $Mode -eq 'Primary') -and $SaveAfterUpdate)
  $readOnly = -not $willSave
  $withWindow = ($Mode -eq 'EditData' -or $Mode -eq 'Primary')
  $presentation = $powerpoint.Presentations.Open($Pptx, $readOnly, $false, $withWindow)
  Mark-Progress 'presentation-open'
  $slide = $presentation.Slides.Item([int]$chartInfo.actual_slide)
  $shape = Get-TargetChartShape $slide ([string]$chartInfo.shape_name)
  if ($shape.HasChart -ne -1) { throw 'Target shape is not a native chart.' }
  $chartObject = $shape.Chart
  $chartData = $chartObject.ChartData
  if (-not [bool]$chartData.IsLinked) { throw 'ChartData.IsLinked is false.' }

  if ($Mode -eq 'Primary') {
    $chartData.Activate()
    $activatedWorkbook = $chartData.Workbook
    $excelApplication = $activatedWorkbook.Application
    $excelPid = Get-OfficePid $excelApplication
    if ($null -eq $excelPid) { throw 'Cannot determine the Excel process created by ChartData.' }
    if ($excelPid -in $pidsBeforeExcel) {
      throw 'ChartData attached to a pre-existing Excel process; verification stopped without closing that process.'
    }
    $excelOwnedByWorker = $true
    Mark-Progress ("excel-pid=" + $excelPid)
    $activateFullName = [string]$activatedWorkbook.FullName
    Mark-Progress ("activate-fullname-raw=" + $activateFullName)
    $activatePath = [IO.Path]::GetFullPath($activateFullName)
    if (-not [string]::Equals($activatePath, $ExpectedWorkbook, [StringComparison]::OrdinalIgnoreCase)) {
      throw "Workbook.FullName mismatch: $activatePath"
    }

    $chartData.ActivateChartDataWindow()
    $excelWorkbook = $chartData.Workbook
    $editDataApplication = $excelWorkbook.Application
    $editDataPid = Get-OfficePid $editDataApplication
    if ($editDataPid -ne $excelPid) {
      throw "ChartDataWindow switched Excel process: $excelPid -> $editDataPid"
    }
    $editDataFullName = [string]$excelWorkbook.FullName
    Mark-Progress ("editdata-fullname-raw=" + $editDataFullName)
    $editDataPath = [IO.Path]::GetFullPath($editDataFullName)
    if (-not [string]::Equals($editDataPath, $ExpectedWorkbook, [StringComparison]::OrdinalIgnoreCase)) {
      throw "Edit Data workbook mismatch: $editDataPath"
    }
    if (-not [bool]$editDataApplication.Visible) { throw 'Excel data window was not visible.' }

    [ordered]@{
      passed = $true
      mode = $Mode
      chart_id = $ChartId
      edit_data = [ordered]@{ passed=$true; workbook=$editDataPath; excel_visible=$true; excel_pid=$excelPid }
      workbook_path_check = [ordered]@{ passed=$true; workbook=$activatePath; excel_pid=$excelPid }
    } | ConvertTo-Json -Depth 8
  } elseif ($Mode -eq 'EditData') {
    $chartData.ActivateChartDataWindow()
    $excelWorkbook = $chartData.Workbook
    $excelApplication = $excelWorkbook.Application
    $excelPid = Get-OfficePid $excelApplication
    if ($null -eq $excelPid) { throw 'Cannot determine the Excel process created by ChartData.' }
    if ($excelPid -in $pidsBeforeExcel) {
      throw 'ChartData attached to a pre-existing Excel process; verification stopped without closing that process.'
    }
    $excelOwnedByWorker = $true
    Mark-Progress ("excel-pid=" + $excelPid)
    $actual = [IO.Path]::GetFullPath([string]$excelWorkbook.FullName)
    if (-not [string]::Equals($actual, $ExpectedWorkbook, [StringComparison]::OrdinalIgnoreCase)) { throw "Edit Data workbook mismatch: $actual" }
    if (-not [bool]$excelApplication.Visible) { throw 'Excel data window was not visible.' }
    [ordered]@{ passed=$true; mode=$Mode; chart_id=$ChartId; workbook=$actual; excel_visible=$true; excel_pid=$excelPid; excel_process_owned_by_worker=$true } | ConvertTo-Json -Depth 6
  } elseif ($Mode -eq 'WorkbookPath') {
    $chartData.Activate()
    $excelWorkbook = $chartData.Workbook
    $excelApplication = $excelWorkbook.Application
    $excelPid = Get-OfficePid $excelApplication
    if ($null -eq $excelPid) { throw 'Cannot determine the Excel process created by ChartData.' }
    if ($excelPid -in $pidsBeforeExcel) {
      throw 'ChartData attached to a pre-existing Excel process; verification stopped without closing that process.'
    }
    $excelOwnedByWorker = $true
    Mark-Progress ("excel-pid=" + $excelPid)
    $actual = [IO.Path]::GetFullPath([string]$excelWorkbook.FullName)
    if (-not [string]::Equals($actual, $ExpectedWorkbook, [StringComparison]::OrdinalIgnoreCase)) { throw "Workbook.FullName mismatch: $actual" }
    [ordered]@{ passed=$true; mode=$Mode; chart_id=$ChartId; workbook=$actual; excel_pid=$excelPid; excel_process_owned_by_worker=$true } | ConvertTo-Json -Depth 6
  } elseif ($Mode -eq 'MutationUpdate') {
    $seriesOrdinal = Resolve-SeriesOrdinal $chartObject $ExpectedSeriesName $SeriesIndex
    Mark-Progress ("series-ordinal=" + $seriesOrdinal + ";series-name=" + $ExpectedSeriesName)
    $before = Get-PointValue $chartObject $seriesOrdinal $PointIndex
    Mark-Progress ("before=" + $before)
    $excelPidsBeforeLinkUpdate = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $shape.LinkFormat.Update()
    Mark-Progress 'link-update-returned'
    $updateExcelPid = Resolve-OwnedUpdateExcelPid $excelPidsBeforeLinkUpdate
    if ($null -ne $updateExcelPid) { Mark-Progress ("update-excel-pid=" + $updateExcelPid) }
    $chartObject.Refresh()
    Mark-Progress 'chart-refresh-returned'
    $deadline = (Get-Date).AddSeconds(45)
    do {
      $after = Get-PointValue $chartObject $seriesOrdinal $PointIndex
      if ([Math]::Abs($after - $ExpectedValue) -le 0.0000001) { break }
      Start-Sleep -Milliseconds 250
      $chartObject.Refresh()
    } while ((Get-Date) -lt $deadline)
    if ([Math]::Abs($after - $ExpectedValue) -gt 0.0000001) { throw "Mutation value mismatch: $after != $ExpectedValue" }
    Mark-Progress ("after=" + $after)
    if ($SaveAfterUpdate) { $presentation.Save(); Mark-Progress 'presentation-saved' }
    [ordered]@{ passed=$true; mode=$Mode; chart_id=$ChartId; before=$before; after=$after; expected=$ExpectedValue; saved=[bool]$SaveAfterUpdate } | ConvertTo-Json -Depth 6
  } else {
    $seriesOrdinal = Resolve-SeriesOrdinal $chartObject $ExpectedSeriesName $SeriesIndex
    $value = Get-PointValue $chartObject $seriesOrdinal $PointIndex
    if (-not [double]::IsNaN($ExpectedValue) -and [Math]::Abs($value - $ExpectedValue) -gt 0.0000001) { throw "Fresh value mismatch: $value != $ExpectedValue" }
    [ordered]@{ passed=$true; mode=$Mode; chart_id=$ChartId; value=$value; expected=$ExpectedValue } | ConvertTo-Json -Depth 6
  }
} finally {
  Release-Com $probeCell; Release-Com $probeSheet
  Release-Com $activatedWorkbook
  Release-Com $editDataApplication
  if ($null -ne $excelWorkbook -and $excelOwnedByWorker) { try { $excelWorkbook.Close($false) } catch {} }
  Release-Com $excelWorkbook
  Release-Com $chartData; Release-Com $chartObject; Release-Com $shape; Release-Com $slide
  if ($null -ne $presentation) { try { if (-not $willSave) { $presentation.Saved = -1 }; $presentation.Close() } catch {}; Release-Com $presentation }
  if ($null -ne $powerpoint -and $null -ne $powerpointPid -and $powerpointPid -notin $pidsBeforePowerPoint) {
    try { $powerpoint.Quit() } catch {}
  }
  Release-Com $powerpoint
  if ($null -ne $excelApplication -and $excelOwnedByWorker) { try { $excelApplication.Quit() } catch {} }
  Release-Com $excelApplication
  [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Start-Sleep -Milliseconds 500
  if ($null -ne $powerpointPid -and $powerpointPid -notin $pidsBeforePowerPoint) {
    Stop-Process -Id $powerpointPid -Force -ErrorAction SilentlyContinue
  }
  if ($null -ne $excelPid -and $excelOwnedByWorker) {
    Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue
  }
  if ($null -ne $updateExcelPid -and $updateExcelPid -notin $pidsBeforeExcel) {
    Stop-Process -Id $updateExcelPid -Force -ErrorAction SilentlyContinue
  }
}
