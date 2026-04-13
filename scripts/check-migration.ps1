param(
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

Write-Host "== Migration check =="
dotnet run --project ".\ConstructionControl.LogicTests\ConstructionControl.LogicTests.csproj" -c $Configuration --no-build

Write-Host "Migration checks passed."
