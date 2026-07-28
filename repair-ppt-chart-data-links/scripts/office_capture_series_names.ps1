# Copyright 2026 XLoffice-Fyl
# SPDX-License-Identifier: Apache-2.0

param(
  [Parameter(Mandatory=$true)][string]$ManifestPath,
  [Parameter(Mandatory=$true)][string]$Pptx,
  [Parameter(Mandatory=$true)][string]$ResultPath
)

$ErrorActionPreference = 'Stop'
$startedAt = Get-Date
$ManifestPath = [IO.Path]::GetFullPath($ManifestPath)
$Pptx = [IO.Path]::GetFullPath($Pptx)
$ResultPath = [IO.Path]::GetFullPath($ResultPath)
$manifest = [IO.File]::ReadAllText($ManifestPath, [Text.Encoding]::UTF8) | ConvertFrom-Json

function Get-LivePowerPointPids {
  @(
    Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object {
      try { $_.Threads.Count -gt 0 } catch { $false }
    } | ForEach-Object { $_.Id }
  )
}

if ((@(Get-LivePowerPointPids)).Count -gt 0) {
  $gate = [ordered]@{
    passed = $false
    requires_user_action = $true
    message = 'PowerPoint is running. Save and close all PowerPoint windows before continuing.'
  }
  [IO.File]::WriteAllText($ResultPath, ($gate | ConvertTo-Json -Depth 4), [Text.UTF8Encoding]::new($false))
  $gate | ConvertTo-Json -Depth 4
  exit 20
}

$baselinePids = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$baselineExcelPids = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$powerpoint = $null; $presentation = $null; $powerpointPid = $null

Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class SeriesCaptureWindowPid {
  [DllImport("user32.dll")]
  public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
"@

function Release-Com($value) {
  if ($null -ne $value) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($value) } catch {} }
}

function Get-OfficePid($application) {
  try {
    [uint32]$processId = 0
    [void][SeriesCaptureWindowPid]::GetWindowThreadProcessId([IntPtr][int64]$application.Hwnd, [ref]$processId)
    return [int]$processId
  } catch { return $null }
}

function Stop-OwnedProcess([Nullable[int]]$processId, [int[]]$baseline) {
  if ($null -eq $processId -or $processId -in $baseline) { return }
  if ($null -ne (Get-Process -Id $processId -ErrorAction SilentlyContinue)) {
    Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
  }
}

function Stop-OwnedOfficeDescendants([int]$parentProcessId, [int[]]$baseline) {
  try {
    $children = @(Get-CimInstance Win32_Process -Filter "ParentProcessId=$parentProcessId" -ErrorAction Stop)
    foreach ($child in $children) {
      if ([string]$child.Name -notin @('EXCEL.EXE', 'POWERPNT.EXE')) { continue }
      $childPid = [int]$child.ProcessId
      if ($childPid -in $baseline) { continue }
      Stop-Process -Id $childPid -Force -ErrorAction SilentlyContinue
    }
  } catch {}
}

try {
  $powerpoint = New-Object -ComObject PowerPoint.Application
  $powerpointPid = Get-OfficePid $powerpoint
  if ($null -eq $powerpointPid -or $powerpointPid -in $baselinePids) {
    throw 'Cannot establish an isolated PowerPoint process for final series capture.'
  }
  $powerpoint.Visible = -1
  $powerpoint.DisplayAlerts = 2
  $presentation = $powerpoint.Presentations.Open($Pptx, $true, $false, $false)
  $captures = @()
  foreach ($chart in @($manifest.charts)) {
    $slide = $null; $shape = $null; $chartObject = $null; $collection = $null
    try {
      $slide = $presentation.Slides.Item([int]$chart.actual_slide)
      $shape = $slide.Shapes.Item([string]$chart.shape_name)
      if ($shape.HasChart -ne -1) { throw "Target shape is not a native chart: $($chart.chart_id)" }
      $chartObject = $shape.Chart
      $collection = $chartObject.SeriesCollection()
      $names = @()
      for ($index = 1; $index -le [int]$collection.Count; $index++) {
        $series = $null
        try {
          $series = $collection.Item($index)
          $names += [string]$series.Name
        } finally { Release-Com $series }
      }
      $captures += [ordered]@{
        chart_id = [string]$chart.chart_id
        runtime_series_names = @($names)
      }
    } finally {
      Release-Com $collection; Release-Com $chartObject; Release-Com $shape; Release-Com $slide
    }
  }
  $result = [ordered]@{
    passed = $true
    charts = @($captures)
    powerpoint_pid = $powerpointPid
    powerpoint_launch_count = 1
    elapsed_seconds = [Math]::Round(((Get-Date) - $startedAt).TotalSeconds, 3)
  }
  [IO.File]::WriteAllText($ResultPath, ($result | ConvertTo-Json -Depth 8), [Text.UTF8Encoding]::new($false))
  $result | ConvertTo-Json -Depth 8
}
finally {
  if ($null -ne $presentation) { try { $presentation.Saved = -1; $presentation.Close() } catch {}; Release-Com $presentation }
  if ($null -ne $powerpoint -and $null -ne $powerpointPid -and $powerpointPid -notin $baselinePids) { try { $powerpoint.Quit() } catch {} }
  Release-Com $powerpoint
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Start-Sleep -Milliseconds 350
  Stop-OwnedProcess $powerpointPid $baselinePids
  Stop-OwnedOfficeDescendants $PID @($baselinePids + $baselineExcelPids)
}
