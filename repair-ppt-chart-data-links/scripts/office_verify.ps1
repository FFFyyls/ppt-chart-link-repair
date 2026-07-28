# Copyright 2026 XLoffice-Fyl
# SPDX-License-Identifier: Apache-2.0

param(
  [Parameter(Mandatory=$true)][string]$ManifestPath,
  [string]$PythonPath = 'python'
)

$ErrorActionPreference = 'Stop'
$startedAt = Get-Date
$ManifestPath = [IO.Path]::GetFullPath($ManifestPath)
$scripts = Split-Path -Parent $MyInvocation.MyCommand.Path
$worker = Join-Path $scripts 'office_stage_worker.ps1'
$exporter = Join-Path $scripts 'office_export_slides.ps1'
$builder = Join-Path $scripts 'office_verify_manifest.py'
$mutator = Join-Path $scripts 'workbook_probe.py'
$sourceManifest = [IO.File]::ReadAllText($ManifestPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
$workDir = [IO.Path]::GetFullPath([string]$sourceManifest.work_dir)
$resultPath = Join-Path $workDir 'office-verification.json'
$stageDir = Join-Path $workDir 'office-verification-stages'
$verifyManifestPath = Join-Path $stageDir 'office-verify-manifest.json'
New-Item -ItemType Directory -Force -Path $stageDir | Out-Null
Get-ChildItem -LiteralPath $stageDir -Recurse -Filter '*-progress.log' -File -ErrorAction SilentlyContinue |
  Remove-Item -Force -ErrorAction SilentlyContinue

function Get-LivePowerPointPids {
  @(
    Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object {
      try { $_.Threads.Count -gt 0 } catch { $true }
    } | ForEach-Object { $_.Id }
  )
}

$powerPointPids = @(Get-LivePowerPointPids)
if ($powerPointPids.Count -gt 0) {
  $gate = [ordered]@{
    passed = $false
    requires_user_action = $true
    message = 'PowerPoint正在运行，请保存并关闭所有PowerPoint窗口后继续'
    stage = 'office-verification-gate'
  }
  [IO.File]::WriteAllText($resultPath, ($gate | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false))
  $gate | ConvertTo-Json -Depth 6
  exit 20
}

& $PythonPath $builder --source-manifest $ManifestPath --output $verifyManifestPath
if ($LASTEXITCODE -ne 0) { throw 'Failed to build the Office verification manifest.' }
$manifest = [IO.File]::ReadAllText($verifyManifestPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
$pptx = [IO.Path]::GetFullPath([string]$manifest.output_pptx)
if (-not (Test-Path -LiteralPath $pptx)) { throw "Output presentation not found: $pptx" }

$baselinePowerPointPids = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$baselineExcelPids = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$progressPath = Join-Path $stageDir 'progress.log'
[IO.File]::WriteAllText($progressPath, '', [Text.UTF8Encoding]::new($false))
$allResults = @()
$requiresUserAction = $false

function Mark([string]$message) {
  Add-Content -LiteralPath $progressPath -Encoding UTF8 -Value ((Get-Date).ToString('o') + "`t" + $message)
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
function Assert-ValidWorkbookBackup([string]$path) {
  $item = Get-Item -LiteralPath $path -ErrorAction Stop
  if ($item.Length -le 0) { throw "Workbook backup is empty: $path" }
  $archive = $null
  try {
    $archive = [IO.Compression.ZipFile]::OpenRead($path)
    if ($null -eq $archive.GetEntry('[Content_Types].xml')) {
      throw "Workbook backup is not a valid XLSX package: $path"
    }
  } finally {
    if ($null -ne $archive) { $archive.Dispose() }
  }
}

function Stop-TrackedOffice([string]$workerProgressPath) {
  if (Test-Path -LiteralPath $workerProgressPath) {
    $tracked = @(
      Get-Content -LiteralPath $workerProgressPath -Encoding UTF8 | ForEach-Object {
        if ($_ -match '(?:powerpoint|excel)-pid=(\d+)') { [int]$Matches[1] }
      } | Sort-Object -Unique
    )
    foreach ($processId in $tracked) {
      if ($processId -notin $baselinePowerPointPids -and $processId -notin $baselineExcelPids) {
        Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
      }
    }
  }
  Start-Sleep -Milliseconds 350
}

function Invoke-Worker(
  [string]$mode,
  $chart,
  [string]$probePptx,
  [string]$workbookPath,
  [double]$expectedValue,
  [bool]$saveAfterUpdate
) {
  $probe = $chart.probe
  Mark ("stage:start:{0}:{1}" -f $chart.chart_id, $mode)
  $arguments = @(
    '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $worker,
    '-Mode', $mode,
    '-ManifestPath', $verifyManifestPath,
    '-ChartId', [string]$chart.chart_id,
    '-Pptx', $probePptx,
    '-ExpectedWorkbook', $workbookPath,
    '-ExpectedSeriesName', [string]$probe.series_name,
    '-SeriesIndex', [string][int]$probe.series_index,
    '-PointIndex', [string][int]$probe.point_index
  )
  if (-not [double]::IsNaN($expectedValue)) {
    $arguments += @('-ExpectedValue', $expectedValue.ToString('R', [Globalization.CultureInfo]::InvariantCulture))
  }
  if ($saveAfterUpdate) { $arguments += '-SaveAfterUpdate' }
  $output = & powershell.exe @arguments 2>&1 | Out-String
  $exitCode = $LASTEXITCODE
  $workerProgressPath = Join-Path ([IO.Path]::GetDirectoryName($probePptx)) (([string]$chart.chart_id) + '-' + $mode + '-progress.log')
  Stop-TrackedOffice $workerProgressPath
  if ($exitCode -eq 20) {
    $script:requiresUserAction = $true
    throw "Stage $mode paused because PowerPoint became active: $output"
  }
  if ($exitCode -ne 0) { throw "Stage $mode failed for $($chart.chart_id): $output" }
  $parsed = $output | ConvertFrom-Json
  if (-not [bool]$parsed.passed) { throw "Stage $mode returned passed=false for $($chart.chart_id)" }
  Mark ("stage:passed:{0}:{1}" -f $chart.chart_id, $mode)
  return $parsed
}

foreach ($chart in @($manifest.charts)) {
  $chartId = [string]$chart.chart_id
  $group = @($manifest.groups | Where-Object { @($_.chart_ids) -contains $chartId })[0]
  if ($null -eq $group) { throw "Workbook group not found for $chartId" }
  $probe = $chart.probe
  $workbookPath = [IO.Path]::GetFullPath([string]$group.workbook_path)
  $chartDir = Join-Path $stageDir $chartId
  New-Item -ItemType Directory -Force -Path $chartDir | Out-Null
  $probePptx = Join-Path $chartDir ($chartId + '-probe.pptx')
  $backupWorkbook = Join-Path $chartDir ($chartId + '-workbook-backup.xlsx')
  $backupHashPath = $backupWorkbook + '.sha256'
  $partialBackup = $backupWorkbook + '.partial'
  $partialHash = $backupHashPath + '.partial'
  $chartResult = [ordered]@{
    chart_id = $chartId
    slide = [int]$chart.actual_slide
    workbook_path = $workbookPath
    sheet = [string]$probe.sheet
    cell = [string]$probe.cell
    series_index = [int]$probe.series_index
    series_name = [string]$probe.series_name
    point_index = [int]$probe.point_index
    original_value = [double]$probe.expected_value
    mutated_value = $null
    edit_data = $null
    workbook_path_check = $null
    mutation_update = $null
    fresh_reopen = $null
    workbook_restored_exactly = $false
    user_excel_processes_preserved = $false
    passed = $false
    error = $null
  }
  try {
    Mark ("chart:start:" + $chartId)
    Copy-Item -LiteralPath $pptx -Destination $probePptx -Force
    [IO.File]::SetAttributes($probePptx, [IO.FileAttributes]::Normal)
    $completeBackup = ((Test-Path -LiteralPath $backupWorkbook) -and (Test-Path -LiteralPath $backupHashPath))
    if (-not $completeBackup) {
      Remove-Item -LiteralPath $backupWorkbook,$backupHashPath,$partialBackup,$partialHash -Force -ErrorAction SilentlyContinue
      Copy-Item -LiteralPath $workbookPath -Destination $partialBackup
      $sourceLength = (Get-Item -LiteralPath $workbookPath -ErrorAction Stop).Length
      $backupLength = (Get-Item -LiteralPath $partialBackup -ErrorAction Stop).Length
      if ($sourceLength -ne $backupLength) { throw "Workbook backup length mismatch for $chartId" }
      Assert-ValidWorkbookBackup $partialBackup
      $baselineHash = (Get-FileHash -LiteralPath $partialBackup -Algorithm SHA256).Hash
      [IO.File]::WriteAllText($partialHash, $baselineHash, [Text.UTF8Encoding]::new($false))
      Move-Item -LiteralPath $partialBackup -Destination $backupWorkbook
      Move-Item -LiteralPath $partialHash -Destination $backupHashPath
    } else {
      Assert-ValidWorkbookBackup $backupWorkbook
      $baselineHash = [IO.File]::ReadAllText($backupHashPath, [Text.Encoding]::UTF8).Trim()
      $actualBackupHash = (Get-FileHash -LiteralPath $backupWorkbook -Algorithm SHA256).Hash
      if (-not [string]::Equals($actualBackupHash, $baselineHash, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Workbook backup hash mismatch for $chartId"
      }
    }
    # Always start from the committed backup, including retries after interruption.
    Copy-Item -LiteralPath $backupWorkbook -Destination $workbookPath -Force
    $original = [double]$probe.expected_value
    $mutated = $original + [Math]::Max([Math]::Abs($original) * 0.137, 0.123456789)
    $chartResult.mutated_value = $mutated
    $primary = Invoke-Worker 'Primary' $chart $probePptx $workbookPath ([double]::NaN) $false
    $chartResult.edit_data = $primary.edit_data
    $chartResult.workbook_path_check = $primary.workbook_path_check
    Mark ("excel:mutate:{0}:{1}!{2}" -f $chartId, $probe.sheet, $probe.cell)
    & $PythonPath $mutator --workbook $workbookPath --sheet ([string]$probe.sheet) --cell ([string]$probe.cell) --value ($mutated.ToString('R', [Globalization.CultureInfo]::InvariantCulture)) | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "Workbook mutation failed for $chartId" }
    $chartResult.mutation_update = Invoke-Worker 'MutationUpdate' $chart $probePptx $workbookPath $mutated $true
    $chartResult.fresh_reopen = Invoke-Worker 'ReadValue' $chart $probePptx $workbookPath $mutated $false
    $chartResult.passed = $true
    Mark ("chart:passed:" + $chartId)
  } catch {
    $chartResult.error = $_.Exception.ToString()
    $chartResult.passed = $false
    Mark ("chart:failed:" + $chartId)
  } finally {
    Get-ChildItem -LiteralPath $chartDir -Filter '*-progress.log' -File -ErrorAction SilentlyContinue |
      ForEach-Object { Stop-TrackedOffice $_.FullName }
    Remove-Item -LiteralPath $partialBackup,$partialHash -Force -ErrorAction SilentlyContinue
    if ((Test-Path -LiteralPath $backupWorkbook) -and (Test-Path -LiteralPath $backupHashPath)) {
      try {
        Assert-ValidWorkbookBackup $backupWorkbook
        $journalHash = [IO.File]::ReadAllText($backupHashPath, [Text.Encoding]::UTF8).Trim()
        $actualBackupHash = (Get-FileHash -LiteralPath $backupWorkbook -Algorithm SHA256).Hash
        if (-not [string]::Equals($actualBackupHash, $journalHash, [StringComparison]::OrdinalIgnoreCase)) {
          throw "Workbook backup hash mismatch during restore for $chartId"
        }
        Copy-Item -LiteralPath $backupWorkbook -Destination $workbookPath -Force
        $restoredHash = (Get-FileHash -LiteralPath $workbookPath -Algorithm SHA256).Hash
        $chartResult.workbook_restored_exactly = [string]::Equals($restoredHash, $journalHash, [StringComparison]::OrdinalIgnoreCase)
        if ($chartResult.workbook_restored_exactly) {
          Remove-Item -LiteralPath $backupWorkbook,$backupHashPath -Force
        }
      } catch {
        $chartResult.error = ($chartResult.error + "`n" + $_.Exception.ToString()).Trim()
        $chartResult.passed = $false
      }
    }
    $currentExcelPids = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $chartResult.user_excel_processes_preserved = (@($baselineExcelPids | Where-Object { $_ -notin $currentExcelPids }).Count -eq 0)
    $allResults += [PSCustomObject]$chartResult
  }
  if ($requiresUserAction) { break }
}

$visualResult = [ordered]@{ passed=$false; exports=@(); error=$null }
if (-not $requiresUserAction -and @($allResults | Where-Object { -not $_.passed }).Count -eq 0) {
  $visualDir = Join-Path $stageDir 'visual-exports'
  $visualResultPath = Join-Path $stageDir 'visual-export.json'
  New-Item -ItemType Directory -Force -Path $visualDir | Out-Null
  Get-ChildItem -LiteralPath $visualDir -Filter '*.png' -File -ErrorAction SilentlyContinue |
    Remove-Item -Force -ErrorAction SilentlyContinue
  $visualOutput = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $exporter -ManifestPath $verifyManifestPath -Pptx $pptx -OutputDir $visualDir -ResultPath $visualResultPath 2>&1 | Out-String
  $visualExitCode = $LASTEXITCODE
  if ($visualExitCode -eq 20) { $requiresUserAction = $true }
  if ($visualExitCode -eq 0 -and (Test-Path -LiteralPath $visualResultPath)) {
    $visualResult = [IO.File]::ReadAllText($visualResultPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
  } else {
    $visualResult = [ordered]@{ passed=$false; exports=@(); error=$visualOutput }
  }
}

$workerProgress = @(
  Get-ChildItem -LiteralPath $stageDir -Recurse -Filter '*-progress.log' -File -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -ne 'progress.log' -and $_.Name -ne 'visual-export-progress.log' } |
    ForEach-Object { Get-Content -LiteralPath $_.FullName -Encoding UTF8 }
)
$workerPowerPointLaunches = @($workerProgress | Where-Object { $_ -match 'powerpoint-pid=' }).Count
$chartDataExcelLaunches = @($workerProgress | Where-Object { $_ -match 'excel-pid=' -and $_ -notmatch 'update-excel-pid=' }).Count
$linkUpdateExcelLaunches = @($workerProgress | Where-Object { $_ -match 'update-excel-pid=' }).Count
$officeLaunchRecords = [ordered]@{
  powerpoint = $workerPowerPointLaunches + $(if ($visualResult.passed) { 1 } else { 0 })
  chartdata_excel = $chartDataExcelLaunches
  link_update_excel = $linkUpdateExcelLaunches
}

$summary = [ordered]@{
  schema_version = '2.0'
  charts = @($allResults)
  total = @($allResults).Count
  passed_count = @($allResults | Where-Object { $_.passed -and $_.workbook_restored_exactly -and $_.user_excel_processes_preserved }).Count
  failed_count = @($allResults | Where-Object { -not $_.passed -or -not $_.workbook_restored_exactly -or -not $_.user_excel_processes_preserved }).Count
  passed = (@($allResults | Where-Object { -not $_.passed -or -not $_.workbook_restored_exactly -or -not $_.user_excel_processes_preserved }).Count -eq 0 -and [bool]$visualResult.passed)
  requires_user_action = $requiresUserAction
  delivery_pptx_unchanged_by_mutation_probe = $true
  visual_exports = @($visualResult.exports)
  visual_export = $visualResult
  office_launch_records = $officeLaunchRecords
  elapsed_seconds = [Math]::Round(((Get-Date) - $startedAt).TotalSeconds, 3)
}
[IO.File]::WriteAllText($resultPath, ($summary | ConvertTo-Json -Depth 16), [Text.UTF8Encoding]::new($false))
$summary | ConvertTo-Json -Depth 16
if ($summary.requires_user_action) { exit 20 }
if (-not $summary.passed) { exit 1 }
