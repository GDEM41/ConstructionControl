$ErrorActionPreference = "Stop"

$desktopPath = [Environment]::GetFolderPath([Environment+SpecialFolder]::DesktopDirectory)
$packagesRoot = Join-Path $PSScriptRoot "packages"

$bundledPackages = @(
    @{
        Name = "SoftMaker Office Professional"
        ZipPath = Join-Path $desktopPath "SoftMaker.Office.Professional.v2024.1230.1206.zip"
        ExpandedFolder = "SoftMaker.Office.Professional.v2024.1230.1206"
        VerificationFile = "PORTABLE.cmd"
    },
    @{
        Name = "PDF-XChange PRO"
        ZipPath = Join-Path $desktopPath "PDF-XChange.PRO.v10.8.4.409.zip"
        ExpandedFolder = "PDF-XChange.PRO.v10.8.4.409"
        VerificationFile = "INSTALL.cmd"
    }
)

if (-not (Test-Path -LiteralPath $packagesRoot)) {
    New-Item -ItemType Directory -Path $packagesRoot | Out-Null
}

foreach ($package in $bundledPackages) {
    if (-not (Test-Path -LiteralPath $package.ZipPath)) {
        throw "Package archive not found for '$($package.Name)': $($package.ZipPath)"
    }

    $targetDir = Join-Path $packagesRoot $package.ExpandedFolder
    if (Test-Path -LiteralPath $targetDir) {
        Remove-Item -LiteralPath $targetDir -Recurse -Force
    }

    Write-Host "Preparing package: $($package.Name)"
    Expand-Archive -LiteralPath $package.ZipPath -DestinationPath $packagesRoot -Force

    $verificationPath = Join-Path $targetDir $package.VerificationFile
    if (-not (Test-Path -LiteralPath $verificationPath)) {
        throw "Package '$($package.Name)' was expanded, but '$($package.VerificationFile)' was not found in $targetDir"
    }
}

$isccCandidates = @(
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
    "${env:ProgramFiles}\Inno Setup 6\ISCC.exe"
)

$iscc = $isccCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) {
    Write-Error "Inno Setup 6 was not found. Install it and run this script again."
}

$scriptPath = Join-Path $PSScriptRoot "MasterPRO.iss"
& $iscc $scriptPath
