# Copyright 2026 XLoffice-Fyl
# SPDX-License-Identifier: Apache-2.0

param(
  [Parameter(Mandatory=$true)][string]$ManifestPath,
  [Parameter(Mandatory=$true)][string]$Pptx,
  [Parameter(Mandatory=$true)][string]$OutputDir,
  [Parameter(Mandatory=$true)][string]$ResultPath
)

$ErrorActionPreference = 'Stop'
$startedAt = Get-Date
$manifest = [IO.File]::ReadAllText([IO.Path]::GetFullPath($ManifestPath), [Text.Encoding]::UTF8) | ConvertFrom-Json
$Pptx = [IO.Path]::GetFullPath($Pptx)
$OutputDir = [IO.Path]::GetFullPath($OutputDir)
$ResultPath = [IO.Path]::GetFullPath($ResultPath)
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

function Get-LivePowerPointPids {
  @(
    Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object {
      try { $_.Threads.Count -gt 0 } catch { $false }
    } | ForEach-Object { $_.Id }
  )
}

if ((@(Get-LivePowerPointPids)).Count -gt 0) {
  [ordered]@{
    passed = $false
    requires_user_action = $true
    message = 'PowerPoint is running. Save and close all PowerPoint windows before continuing.'
  } | ConvertTo-Json -Depth 4
  exit 20
}

$baselinePids = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$powerpoint = $null; $presentation = $null; $powerpointPid = $null
$progressPath = Join-Path ([IO.Path]::GetDirectoryName($ResultPath)) 'visual-export-progress.log'
[IO.File]::WriteAllText($progressPath, '', [Text.UTF8Encoding]::new($false))

Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class VisualExportWindowPid {
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
    [void][VisualExportWindowPid]::GetWindowThreadProcessId([IntPtr][int64]$application.Hwnd, [ref]$processId)
    return [int]$processId
  } catch { return $null }
}

function Stop-OwnedProcess([Nullable[int]]$processId, [int[]]$baseline) {
  if ($null -eq $processId -or $processId -in $baseline) { return }
  if ($null -ne (Get-Process -Id $processId -ErrorAction SilentlyContinue)) {
    Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
  }
}

try {
  $powerpoint = New-Object -ComObject PowerPoint.Application
  $powerpointPid = Get-OfficePid $powerpoint
  if ($null -eq $powerpointPid -or $powerpointPid -in $baselinePids) {
    throw 'Cannot establish an isolated PowerPoint process for slide export.'
  }
  Add-Content -LiteralPath $progressPath -Encoding UTF8 -Value ("powerpoint-pid=" + $powerpointPid)
  $powerpoint.Visible = -1
  $powerpoint.DisplayAlerts = 2
  $presentation = $powerpoint.Presentations.Open($Pptx, $true, $false, $false)
  $slideNumbers = @($manifest.charts | ForEach-Object { [int]$_.actual_slide } | Sort-Object -Unique)
  $exports = @()
  foreach ($slideNumber in $slideNumbers) {
    $slide = $null
    try {
      $slide = $presentation.Slides.Item($slideNumber)
      $path = Join-Path $OutputDir ("slide-{0}.png" -f $slideNumber)
      $slide.Export($path, 'PNG')
      if (-not (Test-Path -LiteralPath $path) -or (Get-Item -LiteralPath $path).Length -le 0) {
        throw "Slide export failed: $slideNumber"
      }
      $exports += [ordered]@{ slide=$slideNumber; path=[IO.Path]::GetFullPath($path) }
    } finally { Release-Com $slide }
  }
  $result = [ordered]@{
    passed = $true
    exports = @($exports)
    unique_slide_count = @($slideNumbers).Count
    powerpoint_pid = $powerpointPid
    powerpoint_launch_count = 1
    elapsed_seconds = [Math]::Round(((Get-Date) - $startedAt).TotalSeconds, 3)
  }
  [IO.File]::WriteAllText($ResultPath, ($result | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false))
  $result | ConvertTo-Json -Depth 6
}
finally {
  if ($null -ne $presentation) { try { $presentation.Saved = -1; $presentation.Close() } catch {}; Release-Com $presentation }
  if ($null -ne $powerpoint -and $null -ne $powerpointPid -and $powerpointPid -notin $baselinePids) { try { $powerpoint.Quit() } catch {} }
  Release-Com $powerpoint
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Start-Sleep -Milliseconds 350
  Stop-OwnedProcess $powerpointPid $baselinePids
}
