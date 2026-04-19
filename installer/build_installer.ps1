$ErrorActionPreference = "Stop"

$desktopPath = [Environment]::GetFolderPath([Environment+SpecialFolder]::DesktopDirectory)
$packagesRoot = Join-Path $PSScriptRoot "packages"
$scriptPath = Join-Path $PSScriptRoot "MasterPRO.iss"
$outputRoot = Join-Path $PSScriptRoot "output"
$outputBaseName = "MasterPRO_Setup"
$buildBaseName = "{0}_{1}" -f $outputBaseName, (Get-Date -Format 'yyyyMMdd_HHmmss')
$buildOutputRoot = Join-Path (Join-Path $env:TEMP "MasterPROInstallerBuild") $buildBaseName

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
    $targetDir = Join-Path $packagesRoot $package.ExpandedFolder
    $verificationPath = Join-Path $targetDir $package.VerificationFile

    if (Test-Path -LiteralPath $verificationPath) {
        Write-Host "Using existing prepared package: $($package.Name)"
    }
    elseif (Test-Path -LiteralPath $package.ZipPath) {
        if (Test-Path -LiteralPath $targetDir) {
            Remove-Item -LiteralPath $targetDir -Recurse -Force
        }

        Write-Host "Preparing package from archive: $($package.Name)"
        Expand-Archive -LiteralPath $package.ZipPath -DestinationPath $packagesRoot -Force
    }
    else {
        throw "Package for '$($package.Name)' was not found. Expected archive '$($package.ZipPath)' or prepared folder '$targetDir'."
    }

    if (-not (Test-Path -LiteralPath $verificationPath)) {
        throw "Package '$($package.Name)' is present, but '$($package.VerificationFile)' was not found in $targetDir"
    }
}

$isccCandidates = @(
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
    "${env:ProgramFiles}\Inno Setup 6\ISCC.exe"
)

$iscc = $isccCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) {
    throw "Inno Setup 6 was not found. Install it and run this script again."
}

if (-not (Test-Path -LiteralPath $outputRoot)) {
    New-Item -ItemType Directory -Path $outputRoot | Out-Null
}

New-Item -ItemType Directory -Path $buildOutputRoot -Force | Out-Null

Get-ChildItem -LiteralPath $outputRoot -Filter '*.tmp' -Force -ErrorAction SilentlyContinue |
    Remove-Item -Force -ErrorAction SilentlyContinue

$buildArgs = @(
    ('/O' + $buildOutputRoot),
    ('/F' + $buildBaseName),
    $scriptPath
)

& $iscc @buildArgs
if ($LASTEXITCODE -ne 0) {
    throw "Inno Setup compilation failed with exit code $LASTEXITCODE."
}

$builtInstaller = Join-Path $buildOutputRoot ($buildBaseName + '.exe')
if (-not (Test-Path -LiteralPath $builtInstaller)) {
    throw "Installer was compiled, but '$builtInstaller' was not found."
}

$finalInstaller = Join-Path $outputRoot ($outputBaseName + '.exe')
try {
    Copy-Item -LiteralPath $builtInstaller -Destination $finalInstaller -Force
    Write-Host "Installer copied to: $finalInstaller"
}
catch {
    Write-Warning "Could not replace '$finalInstaller'. Fresh installer is available at '$builtInstaller'."
}
