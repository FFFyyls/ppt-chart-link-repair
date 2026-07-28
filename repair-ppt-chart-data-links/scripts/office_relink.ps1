# Copyright 2026 XLoffice-Fyl
# SPDX-License-Identifier: Apache-2.0

param(
    [Parameter(Mandatory=$true)][string]$ManifestPath,
    [Parameter(Mandatory=$false)][string]$PythonPath
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)

if (@(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -gt 0) {
    Write-Output '{"passed":false,"requires_user_action":true,"message":"PowerPoint正在运行，请保存并关闭所有PowerPoint窗口后继续"}'
    exit 20
}

function Test-PythonRuntime([string]$Candidate, [string[]]$PrefixArgs) {
    if ([string]::IsNullOrWhiteSpace($Candidate)) { return $false }
    if ($Candidate -match '(?i)[\\/]WindowsApps[\\/]python(?:3)?\.exe$') {
        return $false
    }
    if (-not (Test-Path -LiteralPath $Candidate)) { return $false }
    & $Candidate @PrefixArgs -c "import lxml, openpyxl" *> $null
    return ($LASTEXITCODE -eq 0)
}

$pythonArgs = @()
if (-not [string]::IsNullOrWhiteSpace($PythonPath)) {
    $PythonPath = [IO.Path]::GetFullPath($PythonPath)
    if (-not (Test-PythonRuntime $PythonPath $pythonArgs)) {
        throw "Python runtime is unavailable or missing lxml/openpyxl: $PythonPath"
    }
}
else {
    $candidates = @()
    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($null -ne $py) {
        $candidates += [pscustomobject]@{ Path=$py.Source; Args=@('-3') }
    }
    foreach ($name in @('python3', 'python')) {
        $command = Get-Command $name -ErrorAction SilentlyContinue
        if ($null -ne $command) {
            $candidates += [pscustomobject]@{ Path=$command.Source; Args=@() }
        }
    }
    foreach ($candidate in $candidates) {
        if (Test-PythonRuntime $candidate.Path $candidate.Args) {
            $PythonPath = $candidate.Path
            $pythonArgs = @($candidate.Args)
            break
        }
    }
    if ([string]::IsNullOrWhiteSpace($PythonPath)) {
        throw "No verified Python 3 runtime with lxml and openpyxl was found"
    }
}

$runner = Join-Path $PSScriptRoot 'native_link_pipeline.py'
& $PythonPath @pythonArgs $runner --manifest ([IO.Path]::GetFullPath($ManifestPath))
exit $LASTEXITCODE
