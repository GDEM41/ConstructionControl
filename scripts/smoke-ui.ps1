param(
    [string]$Configuration = "Release",
    [switch]$RunUi
)

$ErrorActionPreference = "Stop"

Write-Host "== Smoke build =="
dotnet build ".\ConstructionControl.csproj" -c $Configuration
dotnet build ".\ConstructionControl.LogicTests\ConstructionControl.LogicTests.csproj" -c $Configuration

Write-Host "== Smoke logic tests =="
dotnet run --project ".\ConstructionControl.LogicTests\ConstructionControl.LogicTests.csproj" -c $Configuration --no-build

if (-not $RunUi) {
    Write-Host "UI launch smoke skipped (use -RunUi to enable)."
    exit 0
}

$exePath = Join-Path $PSScriptRoot "..\bin\$Configuration\net10.0-windows\ConstructionControl.exe"
$exePath = [System.IO.Path]::GetFullPath($exePath)
if (-not (Test-Path $exePath)) {
    throw "Не найден исполняемый файл: $exePath"
}

Write-Host "== UI launch smoke =="
$process = Start-Process -FilePath $exePath -PassThru
Start-Sleep -Seconds 8

if ($process.HasExited) {
    throw "Приложение завершилось раньше времени. Код: $($process.ExitCode)"
}

Stop-Process -Id $process.Id -Force
Write-Host "UI smoke passed."
