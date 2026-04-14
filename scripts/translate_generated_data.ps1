param(
    [string]$RootStatePath = 'C:\Users\kravt\AppData\Local\ConstructionControl\data.json',
    [string]$SecondaryStatePath = 'C:\Users\kravt\AppData\Local\ConstructionControl\Data\data.json'
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$utf8 = New-Object System.Text.UTF8Encoding($false)

function Translate-Name([string]$value) {
    $map = @{
        'РРІР°РЅРѕРІ S.V.' = 'РРІР°РЅРѕРІ РЎ.Р’.'
        'РџРµС‚СЂРѕРІ P.A.' = 'РџРµС‚СЂРѕРІ Рџ.Рђ.'
        'РЎРёРґРѕСЂРѕРІ A.I.' = 'РЎРёРґРѕСЂРѕРІ Рђ.Р.'
        'РЎРјРёСЂРЅРѕРІ O.N.' = 'РЎРјРёСЂРЅРѕРІ Рћ.Рќ.'
        'РљСѓР·РЅРµС†РѕРІ I.M.' = 'РљСѓР·РЅРµС†РѕРІ Р.Рњ.'
        'РћСЂР»РѕРІ A.S.' = 'РћСЂР»РѕРІ Рђ.РЎ.'
        'РњРµР»СЊРЅРёРє D.P.' = 'РњРµР»СЊРЅРёРє Р”.Рџ.'
        'Р СѓРґРµРЅРєРѕ A.I.' = 'Р СѓРґРµРЅРєРѕ Рђ.Р.'
        'Р‘РµР»С‹Р№ K.R.' = 'Р‘РµР»С‹Р№ Рљ.Р .'
        'Р•РіРѕСЂРѕРІ V.S.' = 'Р•РіРѕСЂРѕРІ Р’.РЎ.'
        'Р–СѓРєРѕРІ N.A.' = 'Р–СѓРєРѕРІ Рќ.Рђ.'
        'Р’РѕР»РєРѕРІ R.I.' = 'Р’РѕР»РєРѕРІ Р .Р.'
        'Р¤РµРґРѕСЂРѕРІ D.P.' = 'Р¤РµРґРѕСЂРѕРІ Р”.Рџ.'
        'РўРёС…РѕРЅРѕРІ A.A.' = 'РўРёС…РѕРЅРѕРІ Рђ.Рђ.'
        'Р“РѕСЂР±СѓРЅРѕРІ E.S.' = 'Р“РѕСЂР±СѓРЅРѕРІ Р•.РЎ.'
        'РљР»РёРјРѕРІ M.P.' = 'РљР»РёРјРѕРІ Рњ.Рџ.'
        'Р РѕРјР°РЅРѕРІ A.I.' = 'Р РѕРјР°РЅРѕРІ Рђ.Р.'
        'Р—Р°Р№С†РµРІ D.A.' = 'Р—Р°Р№С†РµРІ Р”.Рђ.'
        'РЎРѕР»РѕРІСЊРµРІ I.R.' = 'РЎРѕР»РѕРІСЊРµРІ Р.Р .'
        'РџРѕР»СЏРєРѕРІ D.S.' = 'РџРѕР»СЏРєРѕРІ Р”.РЎ.'
    }

    if ($null -eq $value) { return $value }
    if ($map.ContainsKey($value)) { return $map[$value] }
    return $value
}

function Translate-Code([string]$value) {
    if ([string]::IsNullOrWhiteSpace($value)) { return $value }
    $value = $value -replace '^PV-', 'РџР’-'
    $value = $value -replace '^PR-', 'РџР -'
    return $value
}

function Translate-Weather([string]$value) {
    if ([string]::IsNullOrWhiteSpace($value)) { return $value }

    if ($value -match '^(?<temp>[+-]?\d+)C\s+(?<kind>.+)$') {
        $kind = switch -Regex ($Matches['kind']) {
            '^clear$' { 'СЏСЃРЅРѕ'; break }
            '^cloudy$' { 'РѕР±Р»Р°С‡РЅРѕ'; break }
            '^rain$' { 'РґРѕР¶РґСЊ'; break }
            '^snow$' { 'СЃРЅРµРі'; break }
            '^fog$' { 'С‚СѓРјР°РЅ'; break }
            default { $Matches['kind'] }
        }
        return ('{0} В°C, {1}' -f $Matches['temp'], $kind)
    }

    if ($value -match '^(?<temp>[+-]?\d+)\s*В°?C\s*(?<kind>РѕР±Р»Р°С‡РЅРѕ|СЏСЃРЅРѕ|РґРѕР¶РґСЊ|СЃРЅРµРі|С‚СѓРјР°РЅ)$') {
        return ('{0} В°C, {1}' -f $Matches['temp'], $Matches['kind'])
    }

    return $value
}

function Translate-Deviation([string]$value) {
    switch ($value) {
        'Axis СЃРјРµРЅР° +3 mm' { return 'РћС‚РєР»РѕРЅРµРЅРёРµ РѕС‚ СЂР°Р·Р±РёРІРѕС‡РЅС‹С… РѕСЃРµР№ +3 РјРј' }
        'Axis СЃРјРµРЅР° +5 mm' { return 'РћС‚РєР»РѕРЅРµРЅРёРµ РѕС‚ СЂР°Р·Р±РёРІРѕС‡РЅС‹С… РѕСЃРµР№ +5 РјРј' }
        'Axis offset +3 mm' { return 'РћС‚РєР»РѕРЅРµРЅРёРµ РѕС‚ СЂР°Р·Р±РёРІРѕС‡РЅС‹С… РѕСЃРµР№ +3 РјРј' }
        'Axis offset +5 mm' { return 'РћС‚РєР»РѕРЅРµРЅРёРµ РѕС‚ СЂР°Р·Р±РёРІРѕС‡РЅС‹С… РѕСЃРµР№ +5 РјРј' }
        'No deviations' { return 'РћС‚РєР»РѕРЅРµРЅРёР№ РЅРµС‚' }
        default { return $value }
    }
}

function Translate-Elements([string]$value) {
    if ([string]::IsNullOrWhiteSpace($value)) { return $value }
    $value = $value -replace '\bPK', 'РџРљ'
    $value = $value -replace '\bPB', 'РџР‘'
    $value = $value -replace '\bRG', 'Р Р“'
    $value = $value -replace '\bRDP', 'Р Р”Рџ'
    $value = $value -replace '\bLM', 'Р›Рњ'
    $value = $value -replace '\bDF', 'Р”Р¤'
    $value = $value -replace '\bK-', 'Рљ-'
    return $value
}

function Translate-InspectionName([string]$value) {
    if ([string]::IsNullOrWhiteSpace($value)) { return $value }

    if ($value -match '^Р•Р¶РµРЅРµРґРµР»СЊРЅРѕ РїСЂРѕРІРµСЂРєР° #(\d+)$') { return ('Р•Р¶РµРЅРµРґРµР»СЊРЅР°СЏ РїСЂРѕРІРµСЂРєР° в„–{0}' -f $Matches[1]) }
    if ($value -match '^Р•Р¶РµРјРµСЃСЏС‡РЅРѕ РїСЂРѕРІРµСЂРєР° #(\d+)$') { return ('Р•Р¶РµРјРµСЃСЏС‡РЅР°СЏ РїСЂРѕРІРµСЂРєР° в„–{0}' -f $Matches[1]) }
    if ($value -match '^Shift РїСЂРѕРІРµСЂРєР° #(\d+)$') { return ('РЎРјРµРЅРЅР°СЏ РїСЂРѕРІРµСЂРєР° в„–{0}' -f $Matches[1]) }
    if ($value -match '^Status control #(\d+)$') { return ('РљРѕРЅС‚СЂРѕР»СЊ СЃРѕСЃС‚РѕСЏРЅРёСЏ в„–{0}' -f $Matches[1]) }

    switch ($value) {
        'Scaffold check all blocks' { return 'РћСЃРјРѕС‚СЂ Р»РµСЃРѕРІ Рё РїРѕРґРјРѕСЃС‚РµР№ РЅР° РІСЃРµС… Р±Р»РѕРєР°С…' }
        'Fence and edge protection control' { return 'РљРѕРЅС‚СЂРѕР»СЊ РѕРіСЂР°Р¶РґРµРЅРёР№ Рё Р·Р°С‰РёС‚С‹ РєСЂРѕРјРѕРє' }
        'Formwork and support check' { return 'РџСЂРѕРІРµСЂРєР° РѕРїР°Р»СѓР±РєРё Рё СЃС‚РѕРµРє' }
        'Helmet and harness check' { return 'РџСЂРѕРІРµСЂРєР° РєР°СЃРѕРє, РїРѕСЏСЃРѕРІ Рё РїСЂРёРІСЏР·РµР№' }
        'Ladders and bridgeways check' { return 'РџСЂРѕРІРµСЂРєР° Р»РµСЃС‚РЅРёС† Рё РїРµСЂРµС…РѕРґРЅС‹С… РјРѕСЃС‚РёРєРѕРІ' }
        'Sling and hook condition check' { return 'РџСЂРѕРІРµСЂРєР° СЃС‚СЂРѕРїРѕРІ, РєСЂСЋРєРѕРІ Рё С‚СЂР°РІРµСЂСЃ' }
        'Welding post and grounding check' { return 'РџСЂРѕРІРµСЂРєР° СЃРІР°СЂРѕС‡РЅРѕРіРѕ РїРѕСЃС‚Р° Рё Р·Р°Р·РµРјР»РµРЅРёСЏ' }
        'Portable tool electrical check' { return 'РџСЂРѕРІРµСЂРєР° РїРµСЂРµРЅРѕСЃРЅРѕРіРѕ СЌР»РµРєС‚СЂРѕРёРЅСЃС‚СЂСѓРјРµРЅС‚Р°' }
        'Pump and hose check' { return 'РџСЂРѕРІРµСЂРєР° Р±РµС‚РѕРЅРѕРЅР°СЃРѕСЃР° Рё СЂСѓРєР°РІРѕРІ' }
        'Fire extinguishers check' { return 'РџСЂРѕРІРµСЂРєР° РѕРіРЅРµС‚СѓС€РёС‚РµР»РµР№ Рё РїРѕР¶Р°СЂРЅС‹С… С‰РёС‚РѕРІ' }
        default { return $value }
    }
}

function Translate-JournalName([string]$value) {
    switch ($value) {
        'PowerTools' { return 'Р­Р»РµРєС‚СЂРѕРёРЅСЃС‚СЂСѓРјРµРЅС‚' }
        'Concrete equipment log' { return 'Р–СѓСЂРЅР°Р» РѕСЃРјРѕС‚СЂР° Р±РµС‚РѕРЅРѕРЅР°СЃРѕСЃР°' }
        'Electrical tool log' { return 'Р–СѓСЂРЅР°Р» РѕСЃРјРѕС‚СЂР° СЌР»РµРєС‚СЂРѕРёРЅСЃС‚СЂСѓРјРµРЅС‚Р°' }
        'Fire safety log' { return 'Р–СѓСЂРЅР°Р» РїРѕР¶Р°СЂРЅРѕР№ Р±РµР·РѕРїР°СЃРЅРѕСЃС‚Рё' }
        'Formwork log' { return 'Р–СѓСЂРЅР°Р» РѕСЃРјРѕС‚СЂР° РѕРїР°Р»СѓР±РєРё' }
        'Ladders log' { return 'Р–СѓСЂРЅР°Р» РѕСЃРјРѕС‚СЂР° Р»РµСЃС‚РЅРёС†' }
        'PPE log' { return 'Р–СѓСЂРЅР°Р» РѕСЃРјРѕС‚СЂР° РЎРР—' }
        'Scaffold inspection log' { return 'Р–СѓСЂРЅР°Р» РѕСЃРјРѕС‚СЂР° Р»РµСЃРѕРІ Рё РїРѕРґРјРѕСЃС‚РµР№' }
        'Temporary fence log' { return 'Р–СѓСЂРЅР°Р» РѕСЃРјРѕС‚СЂР° РІСЂРµРјРµРЅРЅС‹С… РѕРіСЂР°Р¶РґРµРЅРёР№' }
        'Lifting gear log' { return 'Р–СѓСЂРЅР°Р» РѕСЃРјРѕС‚СЂР° РіСЂСѓР·РѕР·Р°С…РІР°С‚РЅС‹С… РїСЂРёСЃРїРѕСЃРѕР±Р»РµРЅРёР№' }
        'Welding station log' { return 'Р–СѓСЂРЅР°Р» РѕСЃРјРѕС‚СЂР° СЃРІР°СЂРѕС‡РЅРѕРіРѕ РїРѕСЃС‚Р°' }
        default { return $value }
    }
}

function Translate-Notes([string]$value) {
    switch ($value) {
        'Need РјР°СЃС‚РµСЂ sign' { return 'РўСЂРµР±СѓРµС‚СЃСЏ РїРѕРґРїРёСЃСЊ РјР°СЃС‚РµСЂР°' }
        'No remarks' { return 'Р—Р°РјРµС‡Р°РЅРёР№ РЅРµС‚' }
        'Fix before end of СЃРјРµРЅР°' { return 'РЈСЃС‚СЂР°РЅРёС‚СЊ РґРѕ РєРѕРЅС†Р° СЃРјРµРЅС‹' }
        'Р§РµРє-Р»РёСЃС‚ complete' { return 'Р§РµРє-Р»РёСЃС‚ Р·Р°РїРѕР»РЅРµРЅ' }
        'Checked, remarks fixed if needed.' { return 'РћСЃРјРѕС‚СЂ РІС‹РїРѕР»РЅРµРЅ, Р·Р°РјРµС‡Р°РЅРёСЏ СѓСЃС‚СЂР°РЅРµРЅС‹ РїСЂРё РЅРµРѕР±С…РѕРґРёРјРѕСЃС‚Рё.' }
        'History record.' { return 'РСЃС‚РѕСЂРёСЏ РїСЂРѕРІРµРґРµРЅРёСЏ РѕСЃРјРѕС‚СЂР°.' }
        default { return $value }
    }
}

function Save-State([string]$path, [object]$state) {
    $json = $state | ConvertTo-Json -Depth 100
    $json = [regex]::Replace(
        $json,
        '\\/Date\(([-+]?\d+)([+-]\d{4})?\)\\/',
        {
            param($match)
            $milliseconds = [int64]$match.Groups[1].Value
            return [DateTimeOffset]::FromUnixTimeMilliseconds($milliseconds).ToLocalTime().ToString(
                'yyyy-MM-ddTHH:mm:sszzz',
                [System.Globalization.CultureInfo]::InvariantCulture)
        })
    [System.IO.File]::WriteAllText($path, $json, $utf8)
}

if (-not (Test-Path $RootStatePath)) {
    throw "РќРµ РЅР°Р№РґРµРЅ С„Р°Р№Р» СЃРѕСЃС‚РѕСЏРЅРёСЏ: $RootStatePath"
}

$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
Copy-Item $RootStatePath ($RootStatePath + ".before_translate_$stamp.bak") -Force
if (Test-Path $SecondaryStatePath) {
    Copy-Item $SecondaryStatePath ($SecondaryStatePath + ".before_translate_$stamp.bak") -Force
}

$state = Get-Content $RootStatePath -Raw -Encoding UTF8 | ConvertFrom-Json

$brigades = @('Р‘СЂРёРіР°РґР° 1', 'Р‘СЂРёРіР°РґР° 2', 'Р‘СЂРёРіР°РґР° 3', 'Р‘СЂРёРіР°РґР° 4')
$brigadierCounter = 0

foreach ($person in @($state.CurrentObject.TimesheetPeople)) {
    $person.FullName = Translate-Name ([string]$person.FullName)
    if ($person.Specialty -eq 'РњРѕРЅС‚Р°Р¶РЅРёРє') {
        $person.Specialty = 'РњРѕРЅС‚Р°Р¶РЅРёРє Р–Р‘Рљ'
    }

    if ($person.IsBrigadier) {
        $person.BrigadeName = $brigades[$brigadierCounter % $brigades.Count]
        $brigadierCounter++
    }
    else {
        $person.BrigadeName = ([string]$person.BrigadeName) -replace '^Brigada\s*', 'Р‘СЂРёРіР°РґР° '
        $person.BrigadeName = Translate-Name ([string]$person.BrigadeName)
    }
}

foreach ($row in @($state.CurrentObject.OtJournal)) {
    $row.FullName = Translate-Name ([string]$row.FullName)
    if ($row.Specialty -eq 'РњРѕРЅС‚Р°Р¶РЅРёРє') { $row.Specialty = 'РњРѕРЅС‚Р°Р¶РЅРёРє Р–Р‘Рљ' }
    if ($row.Profession -eq 'РњРѕРЅС‚Р°Р¶РЅРёРє') { $row.Profession = 'РњРѕРЅС‚Р°Р¶РЅРёРє Р–Р‘Рљ' }
    if ($row.InstructionType -eq 'Primary workplace') { $row.InstructionType = 'РџРµСЂРІРёС‡РЅС‹Р№ РЅР° СЂР°Р±РѕС‡РµРј РјРµСЃС‚Рµ' }
    $row.InstructionNumbers = Translate-Code ([string]$row.InstructionNumbers)
    $row.BrigadierName = Translate-Name ([string]$row.BrigadierName)
}

foreach ($row in @($state.CurrentObject.ProductionJournal)) {
    switch ([string]$row.ActionName) {
        'Montazh' { $row.ActionName = 'РњРѕРЅС‚Р°Р¶' }
        'Kladka' { $row.ActionName = 'РљР»Р°РґРєР°' }
        'Ustroystvo' { $row.ActionName = 'РЈСЃС‚СЂРѕР№СЃС‚РІРѕ' }
    }

    $row.BrigadeName = ([string]$row.BrigadeName) -replace '^Brigada\s*', 'Р‘СЂРёРіР°РґР° '
    $row.Weather = Translate-Weather ([string]$row.Weather)
    $row.Deviations = Translate-Deviation ([string]$row.Deviations)
    $row.ElementsText = Translate-Elements ([string]$row.ElementsText)

    if ([string]$row.WorkKey -match '^(Montazh|Kladka|Ustroystvo)::(.+)$') {
        $action = switch ($Matches[1]) {
            'Montazh' { 'РњРѕРЅС‚Р°Р¶' }
            'Kladka' { 'РљР»Р°РґРєР°' }
            'Ustroystvo' { 'РЈСЃС‚СЂРѕР№СЃС‚РІРѕ' }
        }
        $row.WorkKey = $action + '::' + $Matches[2]
    }
}

foreach ($row in @($state.CurrentObject.InspectionJournal)) {
    $row.JournalName = Translate-JournalName ([string]$row.JournalName)
    $row.JournalDisplay = Translate-JournalName ([string]$row.JournalDisplay)
    $row.InspectionName = Translate-InspectionName ([string]$row.InspectionName)
    $row.InspectionDisplay = Translate-InspectionName ([string]$row.InspectionDisplay)
    $row.Notes = Translate-Notes ([string]$row.Notes)
    $row.NotesDisplay = Translate-Notes ([string]$row.NotesDisplay)
}

Save-State $RootStatePath $state
if (Test-Path (Split-Path $SecondaryStatePath -Parent)) {
    Copy-Item $RootStatePath $SecondaryStatePath -Force
}

$check = Get-Content $RootStatePath -Raw -Encoding UTF8 | ConvertFrom-Json
Write-Host 'ROOT_OK'
$check.CurrentObject.TimesheetPeople | Select-Object -First 8 FullName, Specialty, BrigadeName, IsBrigadier | Format-Table -AutoSize
Write-Host ''
$check.CurrentObject.ProductionJournal | Select-Object -First 8 ActionName, BrigadeName, Weather, Deviations, ElementsText | Format-Table -Wrap -AutoSize
Write-Host ''
$check.CurrentObject.InspectionJournal | Select-Object -First 8 JournalName, InspectionName, Notes | Format-Table -Wrap -AutoSize
