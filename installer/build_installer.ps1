$ErrorActionPreference = "Stop"

$isccCandidates = @(
    "$env:ProgramFiles(x86)\Inno Setup 6\ISCC.exe",
    "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
)

$iscc = $isccCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) {
    Write-Error "Inno Setup 6 не найден. Установите Inno Setup и повторите."
}

$scriptPath = Join-Path $PSScriptRoot "MasterPRO.iss"
& $iscc $scriptPath
