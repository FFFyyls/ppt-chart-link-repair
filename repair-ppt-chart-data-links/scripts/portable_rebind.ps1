# Copyright 2026 XLoffice-Fyl
# SPDX-License-Identifier: Apache-2.0

param(
  [string]$ManifestPath = (Join-Path $PSScriptRoot 'bundle-manifest.json'),
  [switch]$NoOpenAfterSuccess
)

$ErrorActionPreference = 'Stop'
$ManifestPath = [IO.Path]::GetFullPath($ManifestPath)
$bundleRoot = Split-Path -Parent $ManifestPath
$manifest = [IO.File]::ReadAllText($ManifestPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
$resultPath = Join-Path $bundleRoot '本机重绑测试报告.json'
$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('ppt-chart-portable-' + [Guid]::NewGuid().ToString('N'))
$baselinePowerPointPids = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$baselineExcelPids = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$ownedOfficePids = [Collections.Generic.HashSet[int]]::new()
$openAfterSuccess = $null

function Get-LivePowerPointPids {
  @(
    Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object {
      try { $_.Threads.Count -gt 0 } catch { $true }
    } | ForEach-Object { $_.Id }
  )
}

if (@(Get-LivePowerPointPids).Count -gt 0) {
  $gate = [ordered]@{
    passed = $false
    requires_user_action = $true
    message = 'PowerPoint正在运行，请保存并关闭所有PowerPoint窗口后继续'
  }
  [IO.File]::WriteAllText($resultPath, ($gate | ConvertTo-Json -Depth 5), [Text.UTF8Encoding]::new($false))
  Write-Host $gate.message -ForegroundColor Yellow
  exit 20
}

if ($bundleRoot.StartsWith('\\')) { throw '不支持 UNC 或网络路径，请先复制到本地磁盘。' }
$drive = Get-PSDrive -Name ([IO.Path]::GetPathRoot($bundleRoot).Substring(0,1)) -ErrorAction SilentlyContinue
if ($null -ne $drive -and $drive.DisplayRoot) { throw '不支持映射网络驱动器，请先复制到本地磁盘。' }
if ($bundleRoot -match '(?i)[\\/]OneDrive([\\/]|$)' -or $bundleRoot -match '(?i)[\\/]SharePoint([\\/]|$)') {
  throw '暂不支持 OneDrive 或 SharePoint 路径，请先复制到普通本地目录。'
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class PortableOfficePid {
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
    [void][PortableOfficePid]::GetWindowThreadProcessId([IntPtr][int64]$application.Hwnd, [ref]$processId)
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

function Stop-OwnedOffice {
  foreach ($processId in @($ownedOfficePids)) {
    if ($processId -notin $baselinePowerPointPids -and $processId -notin $baselineExcelPids) {
      Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
    }
  }
}

function Get-Sha256([string]$path) {
  return (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Convert-ToOfficeTarget([string]$path) {
  $builder = New-Object Text.StringBuilder
  foreach ($character in [IO.Path]::GetFullPath($path).ToCharArray()) {
    if ($character -eq ' ') { [void]$builder.Append('%20') }
    elseif ($character -eq '#') { [void]$builder.Append('%23') }
    elseif ($character -eq '%') { [void]$builder.Append('%25') }
    elseif ($character -eq '?') { [void]$builder.Append('%3F') }
    elseif ([int]$character -lt 32) { [void]$builder.Append(('%{0:X2}' -f [int]$character)) }
    else { [void]$builder.Append($character) }
  }
  return 'file:///' + $builder.ToString()
}

function Set-PptLinkTarget([string]$pptx, [string]$relationshipPart, [string]$workbookPath) {
  $archive = $null; $entry = $null; $reader = $null; $writer = $null
  try {
    $archive = [IO.Compression.ZipFile]::Open($pptx, [IO.Compression.ZipArchiveMode]::Update)
    $entry = $archive.GetEntry($relationshipPart)
    if ($null -eq $entry) { throw "PPT relationship part not found: $relationshipPart" }
    $reader = New-Object IO.StreamReader($entry.Open(), [Text.Encoding]::UTF8, $true)
    $text = $reader.ReadToEnd()
    $reader.Close(); $reader = $null
    [xml]$xml = $text
    $targetRelations = @($xml.Relationships.Relationship | Where-Object {
      ([string]$_.Type).EndsWith('/oleObject') -and [string]$_.TargetMode -eq 'External'
    })
    if ($targetRelations.Count -ne 1) {
      throw "Expected one external chart relationship in $relationshipPart; found $($targetRelations.Count)"
    }
    $targetRelations[0].Target = Convert-ToOfficeTarget $workbookPath
    $settings = New-Object Xml.XmlWriterSettings
    $settings.Encoding = [Text.UTF8Encoding]::new($false)
    $settings.Indent = $false
    $buffer = New-Object IO.MemoryStream
    $xmlWriter = [Xml.XmlWriter]::Create($buffer, $settings)
    $xml.Save($xmlWriter); $xmlWriter.Close()
    $bytes = $buffer.ToArray(); $buffer.Dispose()
    $entry.Delete(); $entry = $null
    $entry = $archive.CreateEntry($relationshipPart, [IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    $stream.Write($bytes, 0, $bytes.Length)
    $stream.Close()
  } finally {
    if ($null -ne $reader) { $reader.Dispose() }
    if ($null -ne $writer) { $writer.Dispose() }
    if ($null -ne $archive) { $archive.Dispose() }
  }
}

function Get-PptLinkTarget([string]$pptx, [string]$relationshipPart) {
  $archive = $null; $reader = $null
  try {
    $archive = [IO.Compression.ZipFile]::OpenRead($pptx)
    $entry = $archive.GetEntry($relationshipPart)
    if ($null -eq $entry) { throw "PPT relationship part not found: $relationshipPart" }
    $reader = New-Object IO.StreamReader($entry.Open(), [Text.Encoding]::UTF8, $true)
    [xml]$xml = $reader.ReadToEnd()
    $relations = @($xml.Relationships.Relationship | Where-Object {
      ([string]$_.Type).EndsWith('/oleObject') -and [string]$_.TargetMode -eq 'External'
    })
    if ($relations.Count -ne 1) { throw "External link count mismatch in $relationshipPart" }
    return [string]$relations[0].Target
  } finally {
    if ($null -ne $reader) { $reader.Dispose() }
    if ($null -ne $archive) { $archive.Dispose() }
  }
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
      if ($candidate.HasChart -eq -1 -and $null -ne $suffix -and [string]$candidate.Name -match ("{0}\s*$" -f [regex]::Escape($suffix))) {
        $matches += $index
      }
    } finally { Release-Com $candidate }
  }
  if ($matches.Count -eq 1) { return $slide.Shapes.Item([int]$matches[0]) }
  throw "无法定位第 $($slide.SlideIndex) 张中的图表：$expectedName"
}

function Resolve-SeriesOrdinal($chartObject, [string]$expectedName, [int]$fallbackZeroBased) {
  if ([string]::IsNullOrWhiteSpace($expectedName)) { return $fallbackZeroBased + 1 }
  $collection = $null
  $matches = @()
  try {
    $collection = $chartObject.SeriesCollection()
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
    throw "系列名称无法唯一匹配：$expectedName（匹配数 $($matches.Count)）"
  }
  return [int]$matches[0]
}

function Get-PointValue($chartObject, [int]$seriesOrdinal, [int]$pointIndex) {
  $series = $null
  try {
    $series = $chartObject.SeriesCollection($seriesOrdinal)
    $values = @($series.Values)
    if ($pointIndex -lt 0 -or $pointIndex -ge $values.Count) { throw '测试点超出图表系列范围。' }
    return [double]$values[$pointIndex]
  } finally { Release-Com $series }
}

function Test-PortableChart($chart, [string]$localPptx, [string]$workbookPath, [string]$chartTemp) {
  New-Item -ItemType Directory -Force -Path $chartTemp | Out-Null
  $cloneWorkbook = Join-Path $chartTemp ([IO.Path]::GetFileName($workbookPath))
  $probePptx = Join-Path $chartTemp 'link-probe.pptx'
  Copy-Item -LiteralPath $workbookPath -Destination $cloneWorkbook -Force
  Copy-Item -LiteralPath $localPptx -Destination $probePptx -Force
  Set-PptLinkTarget $probePptx ([string]$chart.chart_rel_part) $cloneWorkbook

  $ppt = $null; $presentation = $null; $slide = $null; $shape = $null; $chartObject = $null
  $chartData = $null; $book = $null; $excel = $null; $sheet = $null; $cell = $null
  $excelPid = $null
  $updateExcelPid = $null
  $excelOwnedByWorker = $false
  $mutated = $null
  try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $pptPid = Get-OfficePid $ppt
    if ($null -eq $pptPid) { throw '无法确定测试 PowerPoint 进程。' }
    [void]$ownedOfficePids.Add([int]$pptPid)
    $ppt.Visible = -1; $ppt.DisplayAlerts = 2
    $presentation = $ppt.Presentations.Open($probePptx, $false, $false, $true)
    $slide = $presentation.Slides.Item([int]$chart.actual_slide)
    $shape = Get-TargetChartShape $slide ([string]$chart.shape_name)
    if ($shape.HasChart -ne -1) { throw '目标对象不是原生 PowerPoint 图表。' }
    $chartObject = $shape.Chart
    $chartData = $chartObject.ChartData
    if (-not [bool]$chartData.IsLinked) { throw 'ChartData.IsLinked 为 false。' }
    $chartData.ActivateChartDataWindow()
    $book = $chartData.Workbook
    $excel = $book.Application
    $excelPid = Get-OfficePid $excel
    if ($null -eq $excelPid) { throw '无法确定编辑数据窗口所属的 Excel 进程。' }
    if ($excelPid -in $baselineExcelPids) {
      throw 'ChartData attached to a pre-existing Excel process; 已停止且不会关闭该进程。'
    }
    $excelOwnedByWorker = $true
    [void]$ownedOfficePids.Add([int]$excelPid)
    $actualWorkbook = [IO.Path]::GetFullPath([string]$book.FullName)
    if (-not [string]::Equals($actualWorkbook, $cloneWorkbook, [StringComparison]::OrdinalIgnoreCase)) {
      throw "编辑数据跳转路径不正确：$actualWorkbook"
    }
    if (-not [bool]$excel.Visible) { throw '编辑数据窗口没有显示。' }
    $probe = $chart.probe
    $sheet = $book.Worksheets.Item([string]$probe.sheet)
    $cell = $sheet.Range([string]$probe.cell)
    $original = [double]$cell.Value2
    $seriesOrdinal = Resolve-SeriesOrdinal $chartObject ([string]$probe.series_name) ([int]$probe.series_index)
    $before = Get-PointValue $chartObject $seriesOrdinal ([int]$probe.point_index)
    if ([Math]::Abs($before - [double]$probe.expected_value) -gt 0.0000001) {
      throw "图表初始值与恢复值不一致：$before"
    }
    $mutated = $original + [Math]::Max([Math]::Abs($original) * 0.137, 0.123456789)
    $cell.Value2 = $mutated
    $book.Save()
    $book.Close($false); Release-Com $book; $book = $null
    $excelPidsBeforeLinkUpdate = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $shape.LinkFormat.Update()
    $updateExcelPid = Resolve-OwnedUpdateExcelPid $excelPidsBeforeLinkUpdate
    if ($null -ne $updateExcelPid) { [void]$ownedOfficePids.Add([int]$updateExcelPid) }
    $chartObject.Refresh()
    $deadline = (Get-Date).AddSeconds(45)
    do {
      $after = Get-PointValue $chartObject $seriesOrdinal ([int]$probe.point_index)
      if ([Math]::Abs($after - $mutated) -le 0.0000001) { break }
      Start-Sleep -Milliseconds 250
      $chartObject.Refresh()
    } while ((Get-Date) -lt $deadline)
    if ([Math]::Abs($after - $mutated) -gt 0.0000001) { throw "Excel 改值后图表没有同步：$after" }
    $presentation.Save()
  } finally {
    Release-Com $cell; Release-Com $sheet
    if ($null -ne $book -and $excelOwnedByWorker) { try { $book.Close($false) } catch {} }
    Release-Com $book
    Release-Com $chartData; Release-Com $chartObject; Release-Com $shape; Release-Com $slide
    if ($null -ne $presentation) { try { $presentation.Close() } catch {}; Release-Com $presentation }
    if ($null -ne $ppt) { try { $ppt.Quit() } catch {}; Release-Com $ppt }
    if ($null -ne $excel -and $excelOwnedByWorker) { try { $excel.Quit() } catch {} }
    Release-Com $excel
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 350
    Stop-OwnedOffice
  }

  $ppt = $null; $presentation = $null; $slide = $null; $shape = $null; $chartObject = $null
  try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $pptPid = Get-OfficePid $ppt
    if ($null -eq $pptPid) { throw '无法确定重开测试 PowerPoint 进程。' }
    [void]$ownedOfficePids.Add([int]$pptPid)
    $ppt.Visible = -1; $ppt.DisplayAlerts = 2
    $presentation = $ppt.Presentations.Open($probePptx, $true, $false, $false)
    $slide = $presentation.Slides.Item([int]$chart.actual_slide)
    $shape = Get-TargetChartShape $slide ([string]$chart.shape_name)
    $chartObject = $shape.Chart
    $seriesOrdinal = Resolve-SeriesOrdinal $chartObject ([string]$chart.probe.series_name) ([int]$chart.probe.series_index)
    $fresh = Get-PointValue $chartObject $seriesOrdinal ([int]$chart.probe.point_index)
    if ([Math]::Abs($fresh - $mutated) -gt 0.0000001) { throw "保存重开后的图表值不正确：$fresh" }
    return [ordered]@{
      passed = $true
      chart_id = [string]$chart.chart_id
      edit_data_window = $true
      workbook_full_name = $cloneWorkbook
      original_value = [double]$chart.probe.expected_value
      mutated_value = $mutated
      fresh_reopen_value = $fresh
      real_data_workbook_was_not_modified = $true
    }
  } finally {
    Release-Com $chartObject; Release-Com $shape; Release-Com $slide
    if ($null -ne $presentation) { try { $presentation.Saved = -1; $presentation.Close() } catch {}; Release-Com $presentation }
    if ($null -ne $ppt) { try { $ppt.Quit() } catch {}; Release-Com $ppt }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 350
    Stop-OwnedOffice
  }
}

$results = @()
try {
  New-Item -ItemType Directory -Force -Path $tempRoot | Out-Null
  $master = Join-Path $bundleRoot ([string]$manifest.master_pptx)
  if (-not (Test-Path -LiteralPath $master)) { throw "找不到可迁移母版：$master" }
  if ((Get-Sha256 $master) -ne ([string]$manifest.master_sha256).ToLowerInvariant()) { throw '可迁移母版哈希校验失败。' }
  foreach ($workbook in @($manifest.workbooks)) {
    $path = Join-Path $bundleRoot ([string]$workbook.relative_path)
    if (-not (Test-Path -LiteralPath $path)) { throw "找不到数据工作簿：$path" }
    if ((Get-Sha256 $path) -ne ([string]$workbook.initial_sha256).ToLowerInvariant()) {
      throw "数据工作簿哈希校验失败：$path"
    }
  }

  $localOutput = Join-Path $bundleRoot ([string]$manifest.local_output_pptx)
  $tempOutput = Join-Path $tempRoot 'local-linked.pptx'
  Copy-Item -LiteralPath $master -Destination $tempOutput -Force
  foreach ($chart in @($manifest.charts)) {
    $workbookPath = [IO.Path]::GetFullPath((Join-Path $bundleRoot ([string]$chart.workbook_relative_path)))
    Set-PptLinkTarget $tempOutput ([string]$chart.chart_rel_part) $workbookPath
    $actualTarget = Get-PptLinkTarget $tempOutput ([string]$chart.chart_rel_part)
    if (-not [string]::Equals($actualTarget, (Convert-ToOfficeTarget $workbookPath), [StringComparison]::OrdinalIgnoreCase)) {
      throw "关系目标写入失败：$($chart.chart_id)"
    }
  }
  Move-Item -LiteralPath $tempOutput -Destination $localOutput -Force

  foreach ($chart in @($manifest.charts)) {
    $workbookPath = [IO.Path]::GetFullPath((Join-Path $bundleRoot ([string]$chart.workbook_relative_path)))
    try {
      $chartResult = Test-PortableChart $chart $localOutput $workbookPath (Join-Path $tempRoot ([string]$chart.chart_id))
      $results += [PSCustomObject]$chartResult
    } catch {
      $results += [PSCustomObject][ordered]@{
        passed = $false
        chart_id = [string]$chart.chart_id
        error = $_.Exception.ToString()
      }
    }
  }
  $missingExcelPids = @($baselineExcelPids | Where-Object { $_ -notin @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) })
  $summary = [ordered]@{
    passed = (@($results | Where-Object { -not $_.passed }).Count -eq 0 -and $missingExcelPids.Count -eq 0)
    requires_user_action = $false
    local_output_pptx = $localOutput
    charts = @($results)
    existing_excel_processes_preserved = ($missingExcelPids.Count -eq 0)
    note = 'Office行为测试使用隔离的Excel克隆和PPT探针；包内真实数据工作簿未被改写。'
  }
  [IO.File]::WriteAllText($resultPath, ($summary | ConvertTo-Json -Depth 12), [Text.UTF8Encoding]::new($false))
  if (-not $summary.passed) { throw '本机重绑未通过全部严格测试，详情见本机重绑测试报告.json。' }
  Write-Host "重绑和测试通过：$localOutput" -ForegroundColor Green
  $openAfterSuccess = $localOutput
} catch {
  if (-not (Test-Path -LiteralPath $resultPath)) {
    $failure = [ordered]@{ passed=$false; requires_user_action=$false; error=$_.Exception.ToString(); charts=@($results) }
    [IO.File]::WriteAllText($resultPath, ($failure | ConvertTo-Json -Depth 12), [Text.UTF8Encoding]::new($false))
  }
  Write-Error $_
  exit 1
} finally {
  Stop-OwnedOffice
  if (Test-Path -LiteralPath $tempRoot) { Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
if ($null -ne $openAfterSuccess -and -not $NoOpenAfterSuccess) { Start-Process -FilePath $openAfterSuccess }
