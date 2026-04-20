param(
    [string]$Configuration = "Release",
    [string]$ProjectPath = (Join-Path $PSScriptRoot "..\ConstructionControl.csproj")
)

$ErrorActionPreference = "Stop"

$desktopPath = [Environment]::GetFolderPath([Environment+SpecialFolder]::DesktopDirectory)
$downloadsPath = Join-Path $env:USERPROFILE "Downloads"
$telegramDownloadsPath = Join-Path $downloadsPath "Telegram Desktop"
$packagesRoot = Join-Path $PSScriptRoot "packages"
$scriptPath = Join-Path $PSScriptRoot "MasterPRO.iss"
$outputRoot = Join-Path $PSScriptRoot "output"
$outputBaseName = "MasterPRO_Setup"
$buildBaseName = "{0}_{1}" -f $outputBaseName, (Get-Date -Format 'yyyyMMdd_HHmmss')
$buildOutputRoot = Join-Path (Join-Path $env:TEMP "MasterPROInstallerBuild") $buildBaseName
$appPublishRoot = Join-Path $buildOutputRoot "app"

$bundledPackages = @(
    @{
        Name = "SoftMaker Office Professional"
        ExpandedFolder = "SoftMaker.Office.Professional.v2024.1230.1206"
        VerificationFile = "PORTABLE.cmd"
        ArchiveCandidates = @(
            (Join-Path $telegramDownloadsPath "SoftMaker.Office.Professional.v2024.1230.1206.zip"),
            (Join-Path $downloadsPath "SoftMaker.Office.Professional.v2024.1230.1206.zip"),
            (Join-Path $desktopPath "SoftMaker.Office.Professional.v2024.1230.1206.zip")
        )
        ExpandedCandidates = @(
            (Join-Path $telegramDownloadsPath "SoftMaker.Office.Professional.v2024.1230.1206"),
            (Join-Path $downloadsPath "SoftMaker.Office.Professional.v2024.1230.1206"),
            (Join-Path $desktopPath "SoftMaker.Office.Professional.v2024.1230.1206")
        )
    },
    @{
        Name = "PDF-XChange PRO"
        ExpandedFolder = "PDF-XChange.PRO.v10.8.4.409"
        VerificationFile = "INSTALL.cmd"
        ArchiveCandidates = @(
            (Join-Path $telegramDownloadsPath "PDF-XChange.PRO.v10.8.4.409.zip"),
            (Join-Path $downloadsPath "PDF-XChange.PRO.v10.8.4.409.zip"),
            (Join-Path $desktopPath "PDF-XChange.PRO.v10.8.4.409.zip")
        )
        ExpandedCandidates = @(
            (Join-Path $telegramDownloadsPath "PDF-XChange.PRO.v10.8.4.409"),
            (Join-Path $downloadsPath "PDF-XChange.PRO.v10.8.4.409"),
            (Join-Path $desktopPath "PDF-XChange.PRO.v10.8.4.409")
        )
    }
)

if (-not (Test-Path -LiteralPath $packagesRoot)) {
    New-Item -ItemType Directory -Path $packagesRoot | Out-Null
}

foreach ($package in $bundledPackages) {
    $targetDir = Join-Path $packagesRoot $package.ExpandedFolder
    $verificationPath = Join-Path $targetDir $package.VerificationFile
    $archivePath = $package.ArchiveCandidates | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
    $expandedSourceDir = $package.ExpandedCandidates | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1

    if (Test-Path -LiteralPath $verificationPath) {
        Write-Host "Using existing prepared package: $($package.Name)"
    }
    elseif ($expandedSourceDir) {
        if (Test-Path -LiteralPath $targetDir) {
            Remove-Item -LiteralPath $targetDir -Recurse -Force
        }

        Write-Host "Preparing package from expanded folder: $($package.Name)"
        Copy-Item -LiteralPath $expandedSourceDir -Destination $targetDir -Recurse -Force
    }
    elseif ($archivePath) {
        if (Test-Path -LiteralPath $targetDir) {
            Remove-Item -LiteralPath $targetDir -Recurse -Force
        }

        Write-Host "Preparing package from archive: $($package.Name)"
        Expand-Archive -LiteralPath $archivePath -DestinationPath $packagesRoot -Force
    }
    else {
        $archiveCandidates = ($package.ArchiveCandidates | Where-Object { $_ }) -join "', '"
        $expandedCandidates = ($package.ExpandedCandidates | Where-Object { $_ }) -join "', '"
        throw "Package for '$($package.Name)' was not found. Expected prepared folder '$targetDir', one of archives '$archiveCandidates', or one of expanded folders '$expandedCandidates'."
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
New-Item -ItemType Directory -Path $appPublishRoot -Force | Out-Null

Get-ChildItem -LiteralPath $outputRoot -Filter '*.tmp' -Force -ErrorAction SilentlyContinue |
    Remove-Item -Force -ErrorAction SilentlyContinue

$resolvedProjectPath = (Resolve-Path -LiteralPath $ProjectPath).Path
Write-Host "Publishing application from: $resolvedProjectPath"

$publishArgs = @(
    "publish",
    $resolvedProjectPath,
    "-c", $Configuration,
    "-o", $appPublishRoot,
    "--nologo"
)

& dotnet @publishArgs
if ($LASTEXITCODE -ne 0) {
    throw "dotnet publish failed with exit code $LASTEXITCODE."
}

$appExePath = Join-Path $appPublishRoot "ConstructionControl.exe"
if (-not (Test-Path -LiteralPath $appExePath)) {
    throw "Published application was not found at '$appExePath'."
}

$versionInfo = (Get-Item -LiteralPath $appExePath).VersionInfo
$appVersion = @(
    $versionInfo.ProductVersion,
    $versionInfo.FileVersion
) |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_ -ne "0.0.0.0" } |
    Select-Object -First 1

if (-not $appVersion) {
    throw "Could not determine application version from '$appExePath'. Make sure assembly version attributes are set."
}

$buildArgs = @(
    ('/O' + $buildOutputRoot),
    ('/F' + $buildBaseName),
    ('/DMyAppVersion=' + $appVersion),
    ('/DMyAppSourceDir=' + $appPublishRoot),
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
