param(
    [string]$StatePath = 'C:\Users\kravt\AppData\Local\ConstructionControl\Data\data.json',
    [int]$Seed = 20260414
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $StatePath)) {
    throw "Р¤Р°Р№Р» СЃРѕСЃС‚РѕСЏРЅРёСЏ РЅРµ РЅР°Р№РґРµРЅ: $StatePath"
}

$utf8 = New-Object System.Text.UTF8Encoding($false)
$jsonText = [System.IO.File]::ReadAllText($StatePath, [System.Text.Encoding]::UTF8)
$state = $jsonText | ConvertFrom-Json
if ($null -eq $state -or $null -eq $state.CurrentObject) {
    throw 'Р’ С„Р°Р№Р»Рµ СЃРѕСЃС‚РѕСЏРЅРёСЏ РѕС‚СЃСѓС‚СЃС‚РІСѓРµС‚ CurrentObject.'
}

$co = $state.CurrentObject
$rng = [System.Random]::new($Seed)
$checkMark = [char]0x2714

function SafeText([object]$value) {
    if ($null -eq $value) { return '' }
    return [string]$value
}

function RandInt([int]$min, [int]$max) {
    if ($max -lt $min) { return $min }
    return $rng.Next($min, $max + 1)
}

function Pick([object[]]$items) {
    if ($null -eq $items -or $items.Count -eq 0) { return $null }
    return $items[$rng.Next(0, $items.Count)]
}

function PickMany([object[]]$items, [int]$count) {
    if ($null -eq $items -or $items.Count -eq 0 -or $count -le 0) { return @() }
    $take = [Math]::Min($count, $items.Count)
    return @($items | Sort-Object { $rng.Next() } | Select-Object -First $take)
}

function NewMap {
    return [ordered]@{}
}

function NewObj([hashtable]$map) {
    return [pscustomobject]$map
}

function Ensure-Property([object]$target, [string]$name, [object]$value) {
    if ($target.PSObject.Properties.Name -contains $name) {
        $target.$name = $value
    }
    else {
        Add-Member -InputObject $target -MemberType NoteProperty -Name $name -Value $value -Force
    }
}

function Get-EntryProperty([object]$entry, [string]$name) {
    if ($null -eq $entry) { return $null }
    if ($entry -is [System.Collections.IDictionary]) {
        if ($entry.Contains($name)) { return $entry[$name] }
        return $null
    }
    $prop = $entry.PSObject.Properties[$name]
    if ($null -ne $prop) { return $prop.Value }
    return $null
}

function Get-DictionaryValue([object]$dictionary, [string]$key) {
    if ($null -eq $dictionary -or [string]::IsNullOrWhiteSpace($key)) { return $null }
    if ($dictionary -is [System.Collections.IDictionary]) {
        if ($dictionary.Contains($key)) { return $dictionary[$key] }
        return $null
    }
    $prop = $dictionary.PSObject.Properties[$key]
    if ($null -ne $prop) { return $prop.Value }
    return $null
}

function Get-PropertyNames([object]$node) {
    if ($null -eq $node) { return @() }
    if ($node -is [System.Collections.IDictionary]) {
        return @($node.Keys | ForEach-Object { [string]$_ })
    }
    return @($node.PSObject.Properties.Name)
}

function Get-ExistingDemandEntry([object]$demandMap, [string]$key) {
    return Get-DictionaryValue $demandMap $key
}

function Get-PreferredUnit([string]$typeName, [string]$materialName, [object]$existingEntry) {
    $existingUnit = (SafeText (Get-EntryProperty $existingEntry 'Unit')).Trim()
    if (-not [string]::IsNullOrWhiteSpace($existingUnit)) { return $existingUnit }

    $joined = ("$typeName $materialName").ToLowerInvariant()
    if ($joined -match 'Р±РµС‚РѕРЅ|СЂР°СЃС‚РІРѕСЂ') { return 'Рј3' }
    if ($joined -match 'РєР°Р±РµР»СЊ|РїСЂРѕРІРѕРґ|С€РЅСѓСЂ') { return 'Рј' }
    if ($joined -match 'РєР»РµР№|С€РїР°С‚Р»РµРІ|РіСЂСѓРЅС‚РѕРІ|РєСЂР°СЃРєР°|СЃРјРµСЃСЊ') { return 'РєРі' }
    return 'С€С‚'
}

if ($null -eq $co.MaterialCatalog -or $co.MaterialCatalog.Count -eq 0) {
    throw 'РљР°С‚Р°Р»РѕРі РјР°С‚РµСЂРёР°Р»РѕРІ РїСѓСЃС‚.'
}

$blockCount = [Math]::Max(1, [int]$co.BlocksCount)
$defaultBlocks = @(1..$blockCount | ForEach-Object { [string]$_ })

$materialsByType = NewMap
foreach ($item in @($co.MaterialCatalog)) {
    if ($null -eq $item) { continue }
    $typeName = (SafeText $item.TypeName).Trim()
    $materialName = (SafeText $item.MaterialName).Trim()
    if ([string]::IsNullOrWhiteSpace($typeName) -or [string]::IsNullOrWhiteSpace($materialName)) { continue }

    if (-not $materialsByType.Contains($typeName)) {
        $materialsByType[$typeName] = New-Object System.Collections.Generic.List[string]
    }

    if (-not $materialsByType[$typeName].Contains($materialName)) {
        $materialsByType[$typeName].Add($materialName)
    }
}

if ($materialsByType.Keys.Count -eq 0) {
    throw 'Р’ РєР°С‚Р°Р»РѕРіРµ РЅРµ РЅР°Р№РґРµРЅРѕ РјР°С‚РµСЂРёР°Р»РѕРІ СЃ С‚РёРїР°РјРё.'
}

$fallbackMarkSets = @(
    @('+0.080', '+3.220', '+6.450'),
    @('0.000', '+3.000', '+6.000'),
    @('0.000', '+3.300', '+6.600')
)

$marksByType = NewMap
$typeIndex = 0
foreach ($typeName in $materialsByType.Keys) {
    $marks = @()

    $summaryMarks = Get-DictionaryValue $co.SummaryMarksByGroup $typeName
    if ($null -ne $summaryMarks) {
        $marks = @($summaryMarks | ForEach-Object { (SafeText $_).Trim() } | Where-Object { $_ -ne '' } | Select-Object -Unique)
    }

    if ($marks.Count -eq 0) {
        $marks = @(
            $co.MaterialCatalog |
            Where-Object { (SafeText $_.TypeName).Trim() -eq $typeName } |
            ForEach-Object { @($_.LevelMarks) } |
            ForEach-Object { (SafeText $_).Trim() } |
            Where-Object { $_ -ne '' } |
            Select-Object -Unique
        )
    }

    if ($marks.Count -eq 0) {
        $marks = @($fallbackMarkSets[$typeIndex % $fallbackMarkSets.Count])
    }

    $marksByType[$typeName] = $marks
    $typeIndex++
}

$oldDemand = $co.Demand
$newDemand = NewMap

foreach ($item in @($co.MaterialCatalog)) {
    if ($null -eq $item) { continue }

    $typeName = (SafeText $item.TypeName).Trim()
    $materialName = (SafeText $item.MaterialName).Trim()
    if ([string]::IsNullOrWhiteSpace($typeName) -or [string]::IsNullOrWhiteSpace($materialName)) { continue }

    $key = "$typeName::$materialName"
    $existingEntry = Get-ExistingDemandEntry $oldDemand $key
    $marks = @($marksByType[$typeName])
    if ($marks.Count -eq 0) { $marks = @('0.000', '+3.000', '+6.000') }

    $blockKeys = @()
    $existingLevels = Get-EntryProperty $existingEntry 'Levels'
    if ($null -ne $existingLevels) {
        $blockKeys = @(Get-PropertyNames $existingLevels | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }
    if ($blockKeys.Count -eq 0) {
        $blockKeys = @($defaultBlocks)
    }

    $levels = NewMap
    $mountedLevels = NewMap
    foreach ($blockKey in $blockKeys) {
        $needRow = NewMap
        $doneRow = NewMap
        foreach ($mark in $marks) {
            $need = RandInt 10 60
            $done = RandInt 0 ([Math]::Max(0, $need - 3))
            $needRow[$mark] = [double]$need
            $doneRow[$mark] = [double]$done
        }
        $levels[$blockKey] = $needRow
        $mountedLevels[$blockKey] = $doneRow
    }

    $item.LevelMarks = @($marks)
    $newDemand[$key] = NewObj @{
        Unit = (Get-PreferredUnit $typeName $materialName $existingEntry)
        Levels = $levels
        MountedLevels = $mountedLevels
        Floors = NewMap
        MountedFloors = NewMap
    }
}

Ensure-Property $co 'Demand' $newDemand

$people = @(
    'РРІР°РЅРѕРІ РЎРµСЂРіРµР№ Р’РёРєС‚РѕСЂРѕРІРёС‡',
    'РџРµС‚СЂРѕРІ РџР°РІРµР» РђРЅРґСЂРµРµРІРёС‡',
    'РЎРёРґРѕСЂРѕРІ РђР»РµРєСЃРµР№ РРіРѕСЂРµРІРёС‡',
    'РЎРјРёСЂРЅРѕРІ РћР»РµРі РќРёРєРѕР»Р°РµРІРёС‡',
    'РљСѓР·РЅРµС†РѕРІ РР»СЊСЏ РњР°РєСЃРёРјРѕРІРёС‡',
    'РћСЂР»РѕРІ РђРЅС‚РѕРЅ РЎРµСЂРіРµРµРІРёС‡',
    'РњРµР»СЊРЅРёРє Р”РјРёС‚СЂРёР№ РџР°РІР»РѕРІРёС‡',
    'Р СѓРґРµРЅРєРѕ РђСЂС‚РµРј РР»СЊРёС‡',
    'Р‘РµР»С‹Р№ РљРёСЂРёР»Р» Р РѕРјР°РЅРѕРІРёС‡',
    'Р•РіРѕСЂРѕРІ Р’Р»Р°РґРёСЃР»Р°РІ РЎРµСЂРіРµРµРІРёС‡',
    'Р–СѓРєРѕРІ РќРёРєРёС‚Р° РђР»РµРєСЃРµРµРІРёС‡',
    'Р’РѕР»РєРѕРІ Р РѕРјР°РЅ РРіРѕСЂРµРІРёС‡',
    'Р¤РµРґРѕСЂРѕРІ Р”Р°РЅРёРёР» РџР°РІР»РѕРІРёС‡',
    'РўРёС…РѕРЅРѕРІ РђР»РµРєСЃРµР№ РђСЂС‚РµРјРѕРІРёС‡',
    'Р“РѕСЂР±СѓРЅРѕРІ Р•РіРѕСЂ РЎРµСЂРіРµРµРІРёС‡',
    'РљР»РёРјРѕРІ РњР°РєСЃРёРј РџР°РІР»РѕРІРёС‡',
    'Р РѕРјР°РЅРѕРІ РђСЂС‚СѓСЂ РРіРѕСЂРµРІРёС‡',
    'Р—Р°Р№С†РµРІ Р”РµРЅРёСЃ РђРЅРґСЂРµРµРІРёС‡',
    'РЎРѕР»РѕРІСЊРµРІ РРіРѕСЂСЊ Р РѕРјР°РЅРѕРІРёС‡',
    'РџРѕР»СЏРєРѕРІ Р”РјРёС‚СЂРёР№ РЎРµСЂРіРµРµРІРёС‡'
)

$specialties = @(
    'РњРѕРЅС‚Р°Р¶РЅРёРє Р–Р‘Рљ',
    'РђСЂРјР°С‚СѓСЂС‰РёРє',
    'Р‘РµС‚РѕРЅС‰РёРє',
    'Р­Р»РµРєС‚СЂРѕРјРѕРЅС‚Р°Р¶РЅРёРє',
    'РЎРІР°СЂС‰РёРє',
    'РљР°РјРµРЅС‰РёРє'
)

$ranks = @('3', '4', '5', '6')
$brigades = @('Р‘СЂРёРіР°РґР° 1', 'Р‘СЂРёРіР°РґР° 2', 'Р‘СЂРёРіР°РґР° 3', 'Р‘СЂРёРіР°РґР° 4')

$instructionByProfession = [ordered]@{
    'РњРѕРЅС‚Р°Р¶РЅРёРє Р–Р‘Рљ' = 'РРћРў-РњР–Р‘Рљ-01, РРћРў-РЎРўР -01'
    'РђСЂРјР°С‚СѓСЂС‰РёРє' = 'РРћРў-РђР Рњ-02, РРћРў-РЎРўР -01'
    'Р‘РµС‚РѕРЅС‰РёРє' = 'РРћРў-Р‘Р•Рў-03, РРћРў-РЎРўР -01'
    'Р­Р»РµРєС‚СЂРѕРјРѕРЅС‚Р°Р¶РЅРёРє' = 'РРћРў-Р­Рњ-04, РРћРў-Р­Р›-01'
    'РЎРІР°СЂС‰РёРє' = 'РРћРў-РЎР’-05, РџР‘-Р“РђР—-01'
    'РљР°РјРµРЅС‰РёРє' = 'РРћРў-РљРђРњ-06, РРћРў-РЎРўР -01'
}
Ensure-Property $co 'OtInstructionNumbersByProfession' $instructionByProfession

$primaryInstructionType = 'РџРµСЂРІРёС‡РЅС‹Р№ РЅР° СЂР°Р±РѕС‡РµРј РјРµСЃС‚Рµ'
$today = Get-Date
$month0 = Get-Date -Year $today.Year -Month $today.Month -Day 1
$monthStarts = @($month0.AddMonths(-2), $month0.AddMonths(-1), $month0, $month0.AddMonths(1))

$timesheetPeople = @()
$otJournal = @()

for ($i = 0; $i -lt $people.Count; $i++) {
    $fullName = $people[$i]
    $personId = [guid]::NewGuid().ToString()
    $specialty = $specialties[$i % $specialties.Count]
    $rank = Pick $ranks
    $brigade = $brigades[$i % $brigades.Count]
    $isBrigadier = (($i % $brigades.Count) -eq 0)
    $dailyHours = if (($i % 5) -eq 0) { 12 } else { 8 }

    $months = @()
    foreach ($monthStart in $monthStarts) {
        $monthKey = $monthStart.ToString('yyyy-MM')
        $daysInMonth = [DateTime]::DaysInMonth($monthStart.Year, $monthStart.Month)
        $isFutureMonth = ($monthStart -eq $month0.AddMonths(1))

        $dayValues = NewMap
        $dayEntries = NewMap

        if (-not $isFutureMonth) {
            foreach ($day in 1..$daysInMonth) {
                $date = [DateTime]::new($monthStart.Year, $monthStart.Month, $day)
                $isWeekend = $date.DayOfWeek -in @([System.DayOfWeek]::Saturday, [System.DayOfWeek]::Sunday)

                $value = ''
                $comment = $null
                $documentAccepted = $null

                if ($isWeekend) {
                    $value = 'Р’'
                }
                else {
                    $roll = RandInt 1 100
                    if ($roll -le 78) {
                        $value = "$dailyHours"
                    }
                    elseif ($roll -le 86) {
                        $value = 'Рќ'
                        $comment = 'РћС‚СЃСѓС‚СЃС‚РІРёРµ Р±РµР· СѓРІР°Р¶РёС‚РµР»СЊРЅРѕР№ РїСЂРёС‡РёРЅС‹'
                        $documentAccepted = $false
                    }
                    elseif ($roll -le 94) {
                        $value = 'Р‘'
                        $comment = 'Р‘РѕР»СЊРЅРёС‡РЅС‹Р№ Р»РёСЃС‚'
                        $documentAccepted = $true
                    }
                    else {
                        $value = 'Рћ'
                        $comment = 'РћС‡РµСЂРµРґРЅРѕР№ РѕС‚РїСѓСЃРє'
                        $documentAccepted = $true
                    }
                }

                $presenceMark = ''
                if ($value -match '^\d+$') {
                    $presenceMark = "$checkMark"
                }

                $dayValues["$day"] = $value
                $dayEntries["$day"] = NewObj @{
                    Value = $value
                    PresenceMark = $presenceMark
                    Comment = $comment
                    DocumentAccepted = $documentAccepted
                    ArrivalMarked = $false
                }
            }
        }

        $months += NewObj @{
            MonthKey = $monthKey
            DayValues = $dayValues
            DayEntries = $dayEntries
        }
    }

    $timesheetPeople += NewObj @{
        PersonId = $personId
        FullName = $fullName
        Specialty = $specialty
        Rank = $rank
        BrigadeName = $brigade
        IsBrigadier = $isBrigadier
        DailyWorkHours = $dailyHours
        Months = @($months)
        ArchivedMonths = @()
    }

    $otJournal += NewObj @{
        PersonId = $personId
        InstructionDate = $today.Date.AddDays(-1 * (RandInt 0 20))
        FullName = $fullName
        Specialty = $specialty
        Rank = $rank
        Profession = $specialty
        InstructionType = $primaryInstructionType
        InstructionNumbers = $instructionByProfession[$specialty]
        RepeatPeriodMonths = 3
        IsBrigadier = $isBrigadier
        BrigadierName = if ($isBrigadier) { $null } else { $brigade }
        IsDismissed = $false
        IsPendingRepeat = $true
        IsRepeatCompleted = $false
        IsScheduledRepeat = $false
    }
}

Ensure-Property $co 'TimesheetPeople' @($timesheetPeople)
Ensure-Property $co 'OtJournal' @($otJournal)

$typeNames = @($materialsByType.Keys)
$deviationsByType = NewMap
foreach ($typeName in $typeNames) {
    $deviationsByType[$typeName] = @(
        'РћС‚РєР»РѕРЅРµРЅРёРµ РѕС‚ СЂР°Р·Р±РёРІРѕС‡РЅС‹С… РѕСЃРµР№ +3 РјРј',
        'РћС‚РєР»РѕРЅРµРЅРёРµ РѕС‚ СЂР°Р·Р±РёРІРѕС‡РЅС‹С… РѕСЃРµР№ +5 РјРј',
        'РћС‚РєР»РѕРЅРµРЅРёР№ РЅРµС‚'
    )
}
Ensure-Property $co 'ProductionDeviationsByType' $deviationsByType

$actions = @('РњРѕРЅС‚Р°Р¶', 'РљР»Р°РґРєР°', 'РЈСЃС‚СЂРѕР№СЃС‚РІРѕ')
$weatherKinds = @('СЏСЃРЅРѕ', 'РѕР±Р»Р°С‡РЅРѕ', 'РґРѕР¶РґСЊ', 'СЃРЅРµРі', 'С‚СѓРјР°РЅ')
$productionJournal = @()

$startDate = $today.Date.AddDays(-60)
for ($offset = 0; $offset -le 60; $offset++) {
    $date = $startDate.AddDays($offset)
    if ($date.DayOfWeek -in @([System.DayOfWeek]::Saturday, [System.DayOfWeek]::Sunday)) { continue }

    $temperature = switch ($date.Month) {
        12 { RandInt -10 2 }
        1 { RandInt -12 1 }
        2 { RandInt -9 3 }
        3 { RandInt -2 8 }
        4 { RandInt 2 16 }
        5 { RandInt 10 24 }
        6 { RandInt 14 28 }
        7 { RandInt 17 31 }
        8 { RandInt 16 30 }
        9 { RandInt 10 22 }
        10 { RandInt 3 14 }
        11 { RandInt -2 8 }
        default { RandInt 0 12 }
    }
    $weather = "$temperature В°C, $(Pick $weatherKinds)"

    $rowsPerDay = RandInt 2 4
    for ($rowIndex = 0; $rowIndex -lt $rowsPerDay; $rowIndex++) {
        $typeName = Pick $typeNames
        $materials = @($materialsByType[$typeName])
        if ($materials.Count -eq 0) { continue }

        $selectedMaterials = @(PickMany $materials (RandInt 1 ([Math]::Min(2, $materials.Count))))
        if ($selectedMaterials.Count -eq 0) { continue }

        $blocksForRow = New-Object System.Collections.Generic.List[string]
        $marksForRow = New-Object System.Collections.Generic.List[string]
        $elementLines = New-Object System.Collections.Generic.List[string]
        $remainingLines = New-Object System.Collections.Generic.List[string]

        foreach ($materialName in $selectedMaterials) {
            $demandEntry = Get-DictionaryValue $co.Demand "$typeName::$materialName"
            if ($null -eq $demandEntry) { continue }

            $levels = Get-EntryProperty $demandEntry 'Levels'
            $mountedLevels = Get-EntryProperty $demandEntry 'MountedLevels'
            $availableBlocks = @(Get-PropertyNames $levels)
            if ($availableBlocks.Count -eq 0) { continue }

            $selectedBlocks = @(PickMany $availableBlocks (RandInt 1 ([Math]::Min(2, $availableBlocks.Count))))
            if ($selectedBlocks.Count -eq 0) { continue }

            $selectedMarks = New-Object System.Collections.Generic.List[string]
            foreach ($blockKey in $selectedBlocks) {
                $levelRow = Get-DictionaryValue $levels $blockKey
                $availableMarks = @(Get-PropertyNames $levelRow)
                if ($availableMarks.Count -eq 0) { continue }
                foreach ($mark in @(PickMany $availableMarks (RandInt 1 ([Math]::Min(2, $availableMarks.Count))))) {
                    if (-not $selectedMarks.Contains($mark)) {
                        $selectedMarks.Add($mark)
                    }
                }
            }
            if ($selectedMarks.Count -eq 0) { continue }

            $availableTotal = 0
            foreach ($blockKey in $selectedBlocks) {
                $levelRow = Get-DictionaryValue $levels $blockKey
                $mountedRow = Get-DictionaryValue $mountedLevels $blockKey
                foreach ($mark in $selectedMarks) {
                    $need = [int][Math]::Floor([double](Get-DictionaryValue $levelRow $mark))
                    $done = [int][Math]::Floor([double](Get-DictionaryValue $mountedRow $mark))
                    $availableTotal += [Math]::Max(0, $need - $done)
                }
            }
            if ($availableTotal -le 0) { continue }

            $quantity = [Math]::Min((RandInt 1 9), $availableTotal)
            if ($quantity -le 0) { continue }

            $leftToAllocate = $quantity
            foreach ($blockKey in $selectedBlocks) {
                if ($leftToAllocate -le 0) { break }
                $levelRow = Get-DictionaryValue $levels $blockKey
                $mountedRow = Get-DictionaryValue $mountedLevels $blockKey
                foreach ($mark in $selectedMarks) {
                    if ($leftToAllocate -le 0) { break }
                    $need = [int][Math]::Floor([double](Get-DictionaryValue $levelRow $mark))
                    $done = [int][Math]::Floor([double](Get-DictionaryValue $mountedRow $mark))
                    $available = [Math]::Max(0, $need - $done)
                    if ($available -le 0) { continue }
                    $take = [Math]::Min($available, $leftToAllocate)
                    $mountedRow[$mark] = [double]($done + $take)
                    $leftToAllocate -= $take
                }
            }

            foreach ($blockKey in $selectedBlocks) {
                if (-not $blocksForRow.Contains($blockKey)) {
                    $blocksForRow.Add($blockKey)
                }
            }
            foreach ($mark in $selectedMarks) {
                if (-not $marksForRow.Contains($mark)) {
                    $marksForRow.Add($mark)
                }
            }

            $elementLines.Add("$materialName - $quantity")

            $remainingBlock = [string](Pick $selectedBlocks)
            $remainingMark = [string](Pick @($selectedMarks))
            $remainingNeed = [int][Math]::Floor([double](Get-DictionaryValue (Get-DictionaryValue $levels $remainingBlock) $remainingMark))
            $remainingDone = [int][Math]::Floor([double](Get-DictionaryValue (Get-DictionaryValue $mountedLevels $remainingBlock) $remainingMark))
            $remainingValue = [Math]::Max(0, $remainingNeed - $remainingDone)
            $remainingLines.Add("$materialName: $remainingBlock $remainingMark вЂ” РѕСЃС‚Р°С‚РѕРє $remainingValue")
        }

        if ($elementLines.Count -eq 0) { continue }

        $productionJournal += NewObj @{
            Date = $date
            ActionName = (Pick $actions)
            WorkName = $typeName
            ElementsText = ($elementLines -join '; ')
            BlocksText = (($blocksForRow | Sort-Object) -join ', ')
            MarksText = (($marksForRow | Sort-Object) -join ', ')
            BrigadeName = (Pick $brigades)
            Weather = $weather
            Deviations = (Pick $deviationsByType[$typeName])
            RequiresHiddenWorkAct = ((RandInt 1 100) -le 30)
            RemainingInfo = ($remainingLines -join [Environment]::NewLine)
            SuppressDateDisplay = ($rowIndex -gt 0)
            SuppressWeatherDisplay = ($rowIndex -gt 0)
            IsAutoCorrectedQuantity = $false
            IsGeneratedCompanion = $false
        }
    }
}

Ensure-Property $co 'ProductionJournal' @($productionJournal | Sort-Object Date, WorkName)

$inspectionTemplates = @(
    (NewObj @{ Journal = 'Р–СѓСЂРЅР°Р» РѕСЃРјРѕС‚СЂР° Р»РµСЃРѕРІ Рё РїРѕРґРјРѕСЃС‚РµР№'; Inspection = 'РћСЃРјРѕС‚СЂ Р»РµСЃРѕРІ Рё РїРѕРґРјРѕСЃС‚РµР№ РЅР° РІСЃРµС… Р±Р»РѕРєР°С…'; Period = 7 }),
    (NewObj @{ Journal = 'Р–СѓСЂРЅР°Р» РѕСЃРјРѕС‚СЂР° РІСЂРµРјРµРЅРЅС‹С… РѕРіСЂР°Р¶РґРµРЅРёР№'; Inspection = 'РљРѕРЅС‚СЂРѕР»СЊ РѕРіСЂР°Р¶РґРµРЅРёР№ Рё Р·Р°С‰РёС‚С‹ РєСЂРѕРјРѕРє'; Period = 7 }),
    (NewObj @{ Journal = 'Р–СѓСЂРЅР°Р» РѕСЃРјРѕС‚СЂР° РіСЂСѓР·РѕР·Р°С…РІР°С‚РЅС‹С… РїСЂРёСЃРїРѕСЃРѕР±Р»РµРЅРёР№'; Inspection = 'РџСЂРѕРІРµСЂРєР° СЃС‚СЂРѕРїРѕРІ, РєСЂСЋРєРѕРІ Рё С‚СЂР°РІРµСЂСЃ'; Period = 7 }),
    (NewObj @{ Journal = 'Р–СѓСЂРЅР°Р» РѕСЃРјРѕС‚СЂР° СЌР»РµРєС‚СЂРѕРёРЅСЃС‚СЂСѓРјРµРЅС‚Р°'; Inspection = 'РџСЂРѕРІРµСЂРєР° РїРµСЂРµРЅРѕСЃРЅРѕРіРѕ СЌР»РµРєС‚СЂРѕРёРЅСЃС‚СЂСѓРјРµРЅС‚Р°'; Period = 10 }),
    (NewObj @{ Journal = 'Р–СѓСЂРЅР°Р» РїРѕР¶Р°СЂРЅРѕР№ Р±РµР·РѕРїР°СЃРЅРѕСЃС‚Рё'; Inspection = 'РџСЂРѕРІРµСЂРєР° РѕРіРЅРµС‚СѓС€РёС‚РµР»РµР№ Рё РїРѕР¶Р°СЂРЅС‹С… С‰РёС‚РѕРІ'; Period = 30 }),
    (NewObj @{ Journal = 'Р–СѓСЂРЅР°Р» РѕСЃРјРѕС‚СЂР° РЎРР—'; Inspection = 'РџСЂРѕРІРµСЂРєР° РєР°СЃРѕРє, РїРѕСЏСЃРѕРІ Рё РїСЂРёРІСЏР·РµР№'; Period = 14 }),
    (NewObj @{ Journal = 'Р–СѓСЂРЅР°Р» РѕСЃРјРѕС‚СЂР° РѕРїР°Р»СѓР±РєРё'; Inspection = 'РџСЂРѕРІРµСЂРєР° РѕРїР°Р»СѓР±РєРё Рё СЃС‚РѕРµРє'; Period = 14 }),
    (NewObj @{ Journal = 'Р–СѓСЂРЅР°Р» РѕСЃРјРѕС‚СЂР° Р±РµС‚РѕРЅРѕРЅР°СЃРѕСЃР°'; Inspection = 'РџСЂРѕРІРµСЂРєР° Р±РµС‚РѕРЅРѕРЅР°СЃРѕСЃР° Рё СЂСѓРєР°РІРѕРІ'; Period = 14 }),
    (NewObj @{ Journal = 'Р–СѓСЂРЅР°Р» РѕСЃРјРѕС‚СЂР° Р»РµСЃС‚РЅРёС†'; Inspection = 'РџСЂРѕРІРµСЂРєР° Р»РµСЃС‚РЅРёС† Рё РїРµСЂРµС…РѕРґРЅС‹С… РјРѕСЃС‚РёРєРѕРІ'; Period = 21 }),
    (NewObj @{ Journal = 'Р–СѓСЂРЅР°Р» РѕСЃРјРѕС‚СЂР° СЃРІР°СЂРѕС‡РЅРѕРіРѕ РїРѕСЃС‚Р°'; Inspection = 'РџСЂРѕРІРµСЂРєР° СЃРІР°СЂРѕС‡РЅРѕРіРѕ РїРѕСЃС‚Р° Рё Р·Р°Р·РµРјР»РµРЅРёСЏ'; Period = 30 })
)

$inspectionJournal = @()
foreach ($template in $inspectionTemplates) {
    $reminderStart = $today.Date.AddDays(-1 * (RandInt 20 120))
    $maxWindow = [Math]::Max(3, [int]$template.Period + 10)
    $deltaDays = RandInt 2 $maxWindow
    $lastDone = $reminderStart.AddDays($deltaDays)
    if ($lastDone -gt $today.Date) {
        $lastDone = $today.Date.AddDays(-1 * (RandInt 0 8))
    }

    $inspectionJournal += NewObj @{
        JournalName = $template.Journal
        InspectionName = $template.Inspection
        ReminderStartDate = $reminderStart
        ReminderPeriodDays = [int]$template.Period
        LastCompletedDate = $lastDone
        Notes = 'РћСЃРјРѕС‚СЂ РІС‹РїРѕР»РЅРµРЅ, Р·Р°РјРµС‡Р°РЅРёСЏ СѓСЃС‚СЂР°РЅРµРЅС‹ РїСЂРё РЅРµРѕР±С…РѕРґРёРјРѕСЃС‚Рё.'
        IsCompletionHistory = $false
    }

    if ((RandInt 1 100) -le 50) {
        $historyDate = $lastDone.AddDays(-1 * (RandInt 7 40))
        $inspectionJournal += NewObj @{
            JournalName = $template.Journal
            InspectionName = $template.Inspection
            ReminderStartDate = $reminderStart
            ReminderPeriodDays = [int]$template.Period
            LastCompletedDate = $historyDate
            Notes = 'РСЃС‚РѕСЂРёСЏ РїСЂРѕРІРµРґРµРЅРёСЏ РѕСЃРјРѕС‚СЂР°.'
            IsCompletionHistory = $true
        }
    }
}

Ensure-Property $co 'InspectionJournal' @($inspectionJournal | Sort-Object JournalName, IsCompletionHistory, LastCompletedDate)

$state.SavedAtUtc = [DateTime]::UtcNow.ToString('O')
$newJson = $state | ConvertTo-Json -Depth 100
$newJson = [regex]::Replace(
    $newJson,
    '\\/Date\(([-+]?\d+)([+-]\d{4})?\)\\/',
    {
        param($match)
        $milliseconds = [int64]$match.Groups[1].Value
        return [DateTimeOffset]::FromUnixTimeMilliseconds($milliseconds).ToLocalTime().ToString(
            'yyyy-MM-ddTHH:mm:sszzz',
            [System.Globalization.CultureInfo]::InvariantCulture)
    })

[System.IO.File]::WriteAllText($StatePath, $newJson, $utf8)

Write-Host "Р“РѕС‚РѕРІРѕ: $StatePath"
Write-Host "РњР°С‚РµСЂРёР°Р»РѕРІ РІ РєР°С‚Р°Р»РѕРіРµ: $($co.MaterialCatalog.Count)"
Write-Host "Р›СЋРґРµР№ РІ С‚Р°Р±РµР»Рµ: $($co.TimesheetPeople.Count)"
Write-Host "РЎС‚СЂРѕРє РІ РћРў: $($co.OtJournal.Count)"
Write-Host "РЎС‚СЂРѕРє РІ РџР : $($co.ProductionJournal.Count)"
Write-Host "РЎС‚СЂРѕРє РІ РѕСЃРјРѕС‚СЂР°С…: $($co.InspectionJournal.Count)"
