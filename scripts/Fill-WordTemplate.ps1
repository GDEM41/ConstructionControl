[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [string]$ReplacementsPath,

    [string]$OutputPath,

    [switch]$Overwrite
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Resolve-ExistingPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PathValue,
        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    if (-not (Test-Path -LiteralPath $PathValue)) {
        throw "$Label not found: $PathValue"
    }

    return [System.IO.Path]::GetFullPath($PathValue)
}

function Convert-Replacements {
    param(
        [Parameter(Mandatory = $true)]
        [object]$RawData
    )

    $result = [ordered]@{}

    if ($RawData -is [System.Collections.IEnumerable] -and -not ($RawData -is [string])) {
        foreach ($item in $RawData) {
            if ($null -eq $item) {
                continue
            }

            $itemProps = $item.PSObject.Properties
            $findProp = $itemProps['find']
            $replaceProp = $itemProps['replace']
            if ($null -eq $findProp -or $null -eq $replaceProp) {
                throw 'For array mode, each replacement item must have "find" and "replace" properties.'
            }

            $findText = [string]$findProp.Value
            $replaceText = [string]$replaceProp.Value
            if ([string]::IsNullOrWhiteSpace($findText)) {
                throw 'Replacement "find" value cannot be empty.'
            }

            $result[$findText] = $replaceText
        }

        if ($result.Count -gt 0) {
            return $result
        }
    }

    foreach ($prop in $RawData.PSObject.Properties) {
        $findText = [string]$prop.Name
        if ([string]::IsNullOrWhiteSpace($findText)) {
            continue
        }

        $result[$findText] = [string]$prop.Value
    }

    if ($result.Count -eq 0) {
        throw 'Replacement file is empty. Use object mode {"find":"replace"} or array mode [{"find":"...","replace":"..."}].'
    }

    return $result
}

function Invoke-ReplaceAll {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Document,
        [Parameter(Mandatory = $true)]
        [string]$FindText,
        [Parameter(Mandatory = $true)]
        [string]$ReplaceText
    )

    # Word Find+ReplaceAll keeps target formatting in place.
    $range = $Document.Content
    $find = $range.Find
    $find.ClearFormatting() | Out-Null
    $find.Replacement.ClearFormatting() | Out-Null
    $null = $find.Execute($FindText, $false, $false, $false, $false, $false, $true, 1, $false, $ReplaceText, 2, $false, $false, $false, $false)
}

$inputFullPath = Resolve-ExistingPath -PathValue $InputPath -Label 'Input document'
$replacementsFullPath = Resolve-ExistingPath -PathValue $ReplacementsPath -Label 'Replacement file'

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $inputDirectory = [System.IO.Path]::GetDirectoryName($inputFullPath)
    $inputName = [System.IO.Path]::GetFileNameWithoutExtension($inputFullPath)
    $inputExt = [System.IO.Path]::GetExtension($inputFullPath)
    $OutputPath = Join-Path $inputDirectory "$inputName - filled$inputExt"
}

$outputFullPath = [System.IO.Path]::GetFullPath($OutputPath)
if ((Test-Path -LiteralPath $outputFullPath) -and -not $Overwrite) {
    throw "Output file already exists: $outputFullPath`nUse -Overwrite to replace it."
}

$replacementsRaw = Get-Content -LiteralPath $replacementsFullPath -Raw -Encoding UTF8
$replacementsData = $replacementsRaw | ConvertFrom-Json
$replacements = Convert-Replacements -RawData $replacementsData

$word = $null
$document = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $document = $word.Documents.Open($inputFullPath, $false, $false)
    $document.TrackRevisions = $false

    foreach ($pair in $replacements.GetEnumerator()) {
        Invoke-ReplaceAll -Document $document -FindText ([string]$pair.Key) -ReplaceText ([string]$pair.Value)
    }

    if ([string]::Equals($inputFullPath, $outputFullPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        $document.Save()
    }
    else {
        $outputDirectory = [System.IO.Path]::GetDirectoryName($outputFullPath)
        if (-not [string]::IsNullOrWhiteSpace($outputDirectory)) {
            New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
        }

        if (Test-Path -LiteralPath $outputFullPath) {
            Remove-Item -LiteralPath $outputFullPath -Force
        }

        $document.SaveAs2($outputFullPath)
    }

    Write-Host "Done. Replacements applied: $($replacements.Count)"
    Write-Host "Output: $outputFullPath"
}
finally {
    if ($document -ne $null) {
        $document.Close([ref]0)
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    }

    if ($word -ne $null) {
        $word.Quit()
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
