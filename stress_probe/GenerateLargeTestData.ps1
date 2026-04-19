$ErrorActionPreference = 'Stop'

$path = Join-Path $PSScriptRoot 'data.json'
$jsonText = [System.IO.File]::ReadAllText($path, [System.Text.UTF8Encoding]::new($false))
$state = $jsonText | ConvertFrom-Json
if (-not $state.CurrentObject) { throw 'CurrentObject not found in data.json' }
$co = $state.CurrentObject

$startDate = [datetime]'2025-04-01'
$endDate = [datetime]'2026-03-31'

function RandInt([int]$min, [int]$max) { Get-Random -Minimum $min -Maximum ($max + 1) }
function Pick([object[]]$items) {
    if (-not $items -or $items.Count -eq 0) { return $null }
    return $items[(Get-Random -Minimum 0 -Maximum $items.Count)]
}
function PickMany([object[]]$items, [int]$min, [int]$max) {
    if (-not $items -or $items.Count -eq 0) { return @() }
    $count = [Math]::Min($items.Count, (RandInt $min $max))
    return $items | Sort-Object { Get-Random } | Select-Object -First $count
}
function Ensure-Property($obj, [string]$name, $value) {
    if (-not $obj.PSObject.Properties[$name]) {
        $obj | Add-Member -NotePropertyName $name -NotePropertyValue $value -Force
    }
}
function OrderedObject([hashtable]$h) { return [pscustomobject]$h }
function FirstString($value, [string]$defaultValue = '') {
    if ($null -eq $value) { return $defaultValue }

    if ($value -is [System.Array]) {
        foreach ($v in $value) {
            if ($null -eq $v) { continue }
            $s = [string]$v
            if (-not [string]::IsNullOrWhiteSpace($s)) { return $s.Trim() }
        }
        return $defaultValue
    }

    $single = [string]$value
    if ([string]::IsNullOrWhiteSpace($single)) { return $defaultValue }
    return $single.Trim()
}

$groupNames = @()
if ($co.MaterialGroups -and @($co.MaterialGroups).Count -gt 0) {
    $groupNames = @($co.MaterialGroups | ForEach-Object { $_.Name } | Where-Object { $_ })
}
if ($groupNames.Count -lt 8 -and $co.MaterialNamesByGroup) {
    $groupNames = @($co.MaterialNamesByGroup.PSObject.Properties.Name)
}
if ($groupNames.Count -lt 8) { throw 'Not enough groups to build a large test dataset.' }

$mainGroups = @($groupNames[0..3])
$extraGroups = @($groupNames[4..7])

$mainCategory = (
    $co.MaterialCatalog |
    Where-Object { @($_.LevelMarks).Count -gt 0 } |
    ForEach-Object { FirstString $_.CategoryName } |
    Where-Object { $_ } |
    Select-Object -Unique |
    Select-Object -First 1
)
if (-not $mainCategory) {
    $mainCategory = (
        $state.Journal |
        ForEach-Object { FirstString $_.Category } |
        Where-Object { $_ } |
        Select-Object -Unique |
        Select-Object -First 1
    )
}
$mainCategory = FirstString $mainCategory 'MAIN'

$extraCategory = (
    $co.MaterialCatalog |
    Where-Object { @($_.LevelMarks).Count -eq 0 } |
    ForEach-Object { FirstString $_.CategoryName } |
    Where-Object { $_ } |
    Select-Object -Unique |
    Select-Object -First 1
)
if (-not $extraCategory) {
    $extraCategory = (
        $state.Journal |
        ForEach-Object { FirstString $_.Category } |
        Select-Object -Unique |
        Where-Object { $_ -and $_ -ne $mainCategory } |
        Select-Object -First 1
    )
}
$extraCategory = FirstString $extraCategory 'EXTRA'

$extraTypes = @(
    $co.MaterialCatalog |
    Where-Object { (FirstString $_.CategoryName) -eq $extraCategory } |
    ForEach-Object { FirstString $_.TypeName } |
    Select-Object -Unique |
    Where-Object { $_ }
)
if ($extraTypes.Count -lt 2) { $extraTypes = @('INTERNAL', 'LOWCOST') }
$extraTypeA = FirstString $extraTypes[0] 'INTERNAL'
$extraTypeB = if ($extraTypes.Count -gt 1) { FirstString $extraTypes[1] $extraTypeA } else { $extraTypeA }

$subcategories = @(
    $state.Journal |
    ForEach-Object { FirstString $_.SubCategory } |
    Select-Object -Unique |
    Where-Object { $_ -and $_.Trim().Length -gt 0 }
)
$extraSubA = if ($subcategories.Count -gt 0) { FirstString $subcategories[0] 'SUB-A' } else { 'SUB-A' }
$extraSubB = if ($subcategories.Count -gt 1) { FirstString $subcategories[1] $extraSubA } else { $extraSubA }

$baseMaterials = @{}
foreach ($g in ($mainGroups + $extraGroups)) {
    $list = @()
    if ($co.MaterialNamesByGroup -and $co.MaterialNamesByGroup.PSObject.Properties[$g]) {
        $list = @($co.MaterialNamesByGroup.$g)
    }
    if ($list.Count -eq 0) {
        $list = @(
            $state.Journal |
            Where-Object { $_.MaterialGroup -eq $g } |
            Select-Object -ExpandProperty MaterialName -Unique
        )
    }
    if ($list.Count -eq 0) { $list = @("MAT-$($g.GetHashCode())-1") }
    $baseMaterials[$g] = @($list | Where-Object { $_ } | Select-Object -Unique)
}

function New-StructuredMaterial([string]$seed) {
    $prefix = ''
    if ($seed -match '^[^\d]+') { $prefix = $Matches[0] } else { $prefix = 'M' }
    if ($seed -match '\.') {
        return "{0}{1}.{2}-{3}" -f $prefix, (RandInt 20 69), (RandInt 10 39), (RandInt 1 60)
    }
    if ($seed -match '-') {
        return "{0}{1}-{2}" -f $prefix, (RandInt 1 80), (RandInt 1 12)
    }
    return "{0}{1}" -f $prefix, (RandInt 100 999)
}

$extendedMaterials = @{}
foreach ($g in $mainGroups) {
    $set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::CurrentCultureIgnoreCase)
    foreach ($m in $baseMaterials[$g]) { [void]$set.Add($m) }
    $seed = $baseMaterials[$g][0]
    while ($set.Count -lt ($baseMaterials[$g].Count + 6)) {
        [void]$set.Add((New-StructuredMaterial $seed))
    }
    $extendedMaterials[$g] = @($set)
}
foreach ($idx in 0..($extraGroups.Count - 1)) {
    $g = $extraGroups[$idx]
    $set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::CurrentCultureIgnoreCase)
    foreach ($m in $baseMaterials[$g]) { [void]$set.Add($m) }
    $n = 1
    while ($set.Count -lt ($baseMaterials[$g].Count + 4)) {
        [void]$set.Add("MAT-$($idx + 1)-$n")
        $n++
    }
    $extendedMaterials[$g] = @($set)
}

Ensure-Property $co 'MaterialNamesByGroup' ([pscustomobject]@{})
foreach ($g in ($mainGroups + $extraGroups)) {
    $co.MaterialNamesByGroup.PSObject.Properties.Remove($g) | Out-Null
    $co.MaterialNamesByGroup | Add-Member -NotePropertyName $g -NotePropertyValue @($extendedMaterials[$g]) -Force
}

$co.MaterialGroups = @(
    foreach ($g in ($mainGroups + $extraGroups)) {
        [pscustomobject]@{
            Name = $g
            Items = @($extendedMaterials[$g])
        }
    }
)

Ensure-Property $co 'SummaryMarksByGroup' ([pscustomobject]@{})
$marksByGroup = @{}
foreach ($g in $mainGroups) {
    $marks = @()
    if ($co.SummaryMarksByGroup.PSObject.Properties[$g]) { $marks = @($co.SummaryMarksByGroup.$g) }
    if ($marks.Count -eq 0) { $marks = @('0.000', '+3.000', '+6.000') }
    if ($marks.Count -lt 4) {
        $last = $marks[-1]
        $num = 0.0
        if ([double]::TryParse(($last -replace '[^0-9\.-]',''), [ref]$num)) {
            $marks += ('{0:+0.000;-0.000;0.000}' -f ($num + 3.0))
        } else {
            $marks += '+9.000'
        }
    }
    $marks = @($marks | Select-Object -Unique)
    $marksByGroup[$g] = $marks
    $co.SummaryMarksByGroup.PSObject.Properties.Remove($g) | Out-Null
    $co.SummaryMarksByGroup | Add-Member -NotePropertyName $g -NotePropertyValue $marks -Force
}
$co.SummaryVisibleGroups = @($mainGroups + $extraGroups)

Ensure-Property $co 'StbByGroup' ([pscustomobject]@{})
Ensure-Property $co 'SupplierByGroup' ([pscustomobject]@{})
$fallbackStb = FirstString ($co.StbByGroup.PSObject.Properties.Value | Select-Object -First 1) 'STD'
$fallbackSupplier = FirstString ($co.SupplierByGroup.PSObject.Properties.Value | Select-Object -First 1) 'SUPPLIER'
foreach ($g in ($mainGroups + $extraGroups)) {
    if (-not $co.StbByGroup.PSObject.Properties[$g]) {
        $val = FirstString $fallbackStb 'STD'
        $co.StbByGroup | Add-Member -NotePropertyName $g -NotePropertyValue $val -Force
    }
    if (-not $co.SupplierByGroup.PSObject.Properties[$g]) {
        $val = FirstString $fallbackSupplier 'SUPPLIER'
        $co.SupplierByGroup | Add-Member -NotePropertyName $g -NotePropertyValue $val -Force
    }
}

$co.MaterialCatalog = @(
    foreach ($g in $mainGroups) {
        foreach ($m in $extendedMaterials[$g]) {
            [pscustomobject]@{
                CategoryName = FirstString $mainCategory 'MAIN'
                TypeName = FirstString $g 'TYPE'
                SubTypeName = ''
                ExtraLevels = @()
                LevelMarks = @($marksByGroup[$g])
                MaterialName = FirstString $m 'M-001'
            }
        }
    }
    foreach ($i in 0..($extraGroups.Count - 1)) {
        $g = $extraGroups[$i]
        $type = if ($i -lt 2) { $extraTypeA } else { $extraTypeB }
        foreach ($m in $extendedMaterials[$g]) {
            [pscustomobject]@{
                CategoryName = FirstString $extraCategory 'EXTRA'
                TypeName = FirstString $type 'INTERNAL'
                SubTypeName = FirstString $g ''
                ExtraLevels = @()
                LevelMarks = @()
                MaterialName = FirstString $m 'M-001'
            }
        }
    }
)

$co.AutoSplitMaterialNames = @(
    foreach ($g in $mainGroups) { $extendedMaterials[$g] }
)
$splitRules = [ordered]@{}
foreach ($m in $co.AutoSplitMaterialNames) {
    $segments = @()
    if ($m -match '^[^\d]+') { $segments += $Matches[0] }
    $segments += ([regex]::Matches($m, '\d+') | ForEach-Object { $_.Value })
    if ($segments.Count -gt 0) { $splitRules[$m] = ($segments -join '|') }
}
$co.MaterialTreeSplitRules = [pscustomobject]$splitRules

$blocksCount = if ($co.BlocksCount -gt 0) { [int]$co.BlocksCount } else { 3 }
$co.BlocksCount = $blocksCount
$co.SameFloorsInBlocks = $true
if (-not $co.FloorsPerBlock -or $co.FloorsPerBlock -lt 3) { $co.FloorsPerBlock = 4 }
$co.HasBasement = $true
$co.FloorsByBlock = [pscustomobject]@{}

$unitByGroup = @{}
foreach ($g in ($mainGroups + $extraGroups)) {
    $u = (
        $state.Journal |
        Where-Object { $_.MaterialGroup -eq $g -and (FirstString $_.Unit) } |
        ForEach-Object { FirstString $_.Unit } |
        Select-Object -Unique |
        Select-Object -First 1
    )
    if (-not $u) { $u = 'pc' }
    $unitByGroup[$g] = FirstString $u 'pc'
}

$demand = [ordered]@{}
foreach ($g in $mainGroups) {
    foreach ($m in $extendedMaterials[$g]) {
        $levels = [ordered]@{}
        $mounted = [ordered]@{}
        foreach ($b in 1..$blocksCount) {
            $lv = [ordered]@{}
            $mv = [ordered]@{}
            foreach ($mark in $marksByGroup[$g]) {
                $need = RandInt 8 34
                $done = RandInt 0 ([Math]::Max(0, $need - 2))
                $lv[$mark] = $need
                $mv[$mark] = $done
            }
            $levels[[string]$b] = [pscustomobject]$lv
            $mounted[[string]$b] = [pscustomobject]$mv
        }
        $demand["$g::$m"] = [pscustomobject]@{
            Unit = $unitByGroup[$g]
            Levels = [pscustomobject]$levels
            MountedLevels = [pscustomobject]$mounted
            Floors = [pscustomobject]@{}
            MountedFloors = [pscustomobject]@{}
        }
    }
}
$co.Demand = [pscustomobject]$demand

$co.ProductionAutoFillSettings = [pscustomobject]@{
    MinQuantityPerRow = 3
    MaxQuantityPerRow = 10
    MinRowsPerRun = 4
    TargetRowsPerRun = 6
    MaxRowsPerRun = 8
    MaxItemsPerRow = 3
    PreferSelectedTypeOnly = $false
    UseBalancedDistribution = $true
    PreferDemandDeficit = $true
    RespectSelectedBlocksAndMarks = $true
    AllowMixedMaterialsInRow = $true
}

Ensure-Property $co 'UiSettings' ([pscustomobject]@{})
Ensure-Property $co.UiSettings 'DisableTree' $false
Ensure-Property $co.UiSettings 'PinTreeByDefault' $false
Ensure-Property $co.UiSettings 'ShowReminderPopup' $true
Ensure-Property $co.UiSettings 'ReminderSnoozeMinutes' 15
Ensure-Property $co.UiSettings 'HideReminderDetails' $false
$co.UiSettings.DisableTree = [bool]$co.UiSettings.DisableTree
$co.UiSettings.PinTreeByDefault = [bool]$co.UiSettings.PinTreeByDefault
$co.UiSettings.ShowReminderPopup = $true
if (-not $co.UiSettings.ReminderSnoozeMinutes -or $co.UiSettings.ReminderSnoozeMinutes -le 0) { $co.UiSettings.ReminderSnoozeMinutes = 15 }
$co.UiSettings.HideReminderDetails = $false

$nameSeed = @($co.TimesheetPeople | Select-Object -ExpandProperty FullName -Unique | Where-Object { $_ })
if ($nameSeed.Count -lt 10) {
    $nameSeed = @(
        'Ivanov Andrey Sergeevich',
        'Petrov Pavel Alekseevich',
        'Sidorov Ilya Viktorovich',
        'Smirnov Roman Nikolaevich',
        'Volkov Oleg Dmitrievich',
        'Egorov Kirill Maksimovich',
        'Kozlov Anton Mihailovich',
        'Fedorov Georgiy Romanovich',
        'Klimov Artur Olegovich',
        'Rudenko Sergey Pavlovich'
    )
}
$specialtySeed = @($co.TimesheetPeople | Select-Object -ExpandProperty Specialty -Unique | Where-Object { $_ })
if ($specialtySeed.Count -eq 0) { $specialtySeed = @('worker', 'fitter', 'welder', 'concrete') }
$rankSeed = @($co.TimesheetPeople | Select-Object -ExpandProperty Rank -Unique | Where-Object { $_ })
if ($rankSeed.Count -eq 0) { $rankSeed = @('4', '5', '6') }
$brigadeSeed = @($co.TimesheetPeople | Select-Object -ExpandProperty BrigadeName -Unique | Where-Object { $_ })
if ($brigadeSeed.Count -lt 4) { $brigadeSeed = 1..8 | ForEach-Object { "Crew-$($_)" } }
$professionSeed = @($co.OtJournal | Select-Object -ExpandProperty Profession -Unique | Where-Object { $_ })
if ($professionSeed.Count -eq 0) { $professionSeed = @('worker') }

$peopleCount = 120
$people = @()
$used = New-Object 'System.Collections.Generic.HashSet[string]'
for ($i = 0; $i -lt $peopleCount; $i++) {
    $base = Pick $nameSeed
    $name = "$base #$($i + 1)"
    while (-not $used.Add($name)) { $name = "$base #$([guid]::NewGuid().ToString().Substring(0,4))" }
    $people += [pscustomobject]@{
        PersonId = [guid]::NewGuid().ToString()
        FullName = $name
        Specialty = (Pick $specialtySeed)
        Rank = (Pick $rankSeed)
        BrigadeName = $brigadeSeed[$i % $brigadeSeed.Count]
        IsBrigadier = $false
    }
}
foreach ($b in $brigadeSeed) {
    $leader = $people | Where-Object BrigadeName -eq $b | Select-Object -First 1
    if ($leader) { $leader.IsBrigadier = $true }
}
$leaders = @{}
foreach ($p in $people | Where-Object IsBrigadier) { $leaders[$p.BrigadeName] = $p.FullName }

$monthKeys = @()
$cursor = Get-Date -Date $startDate
while ($cursor -le $endDate) {
    $monthKeys += $cursor.ToString('yyyy-MM')
    $cursor = $cursor.AddMonths(1)
}

$nonNumericSeed = @(
    $co.TimesheetPeople |
    ForEach-Object {
        $_.Months | ForEach-Object {
            $_.DayValues.PSObject.Properties.Value |
            ForEach-Object { FirstString $_ } |
            Where-Object { $_ -and ($_ -notmatch '^\d') }
        }
    } | Select-Object -Unique
)
if ($nonNumericSeed.Count -eq 0) { $nonNumericSeed = @('OFF', 'ABS', 'SICK', 'VAC', 'TRIP') }
$weekendCode = FirstString ($nonNumericSeed | Select-Object -First 1) 'OFF'
if (-not $weekendCode) { $weekendCode = 'OFF' }
$commentSeed = @(
    $co.TimesheetPeople |
    ForEach-Object {
        $_.Months | ForEach-Object {
            $_.DayEntries.PSObject.Properties.Value.Comment
        }
    } |
    Where-Object { $_ -and $_.Trim().Length -gt 0 } |
    Select-Object -Unique
)
if ($commentSeed.Count -eq 0) { $commentSeed = @('doc provided', 'approved leave', 'business trip') }

$instrTypes = @($co.OtJournal | Select-Object -ExpandProperty InstructionType -Unique | Where-Object { $_ })
if ($instrTypes.Count -lt 2) { $instrTypes = @('PRIMARY', 'REPEAT') }
$primaryType = FirstString $instrTypes[0] 'PRIMARY'
$repeatType = FirstString $instrTypes[1] 'REPEAT'

$timesheetRows = @()
$otRows = @()
foreach ($p in $people) {
    $months = @()
    foreach ($mk in $monthKeys) {
        $y = [int]$mk.Substring(0,4)
        $m = [int]$mk.Substring(5,2)
        $dim = [datetime]::DaysInMonth($y, $m)
        $dayValues = [ordered]@{}
        $dayEntries = [ordered]@{}

        foreach ($day in 1..$dim) {
            $d = [datetime]::new($y, $m, $day)
            $isWeekend = $d.DayOfWeek -in @([System.DayOfWeek]::Saturday, [System.DayOfWeek]::Sunday)

            if ($isWeekend) {
                $value = $weekendCode
            } else {
                $roll = RandInt 1 100
                if ($roll -le 84) { $value = (Pick @('10', '11', '8')) } else { $value = (Pick $nonNumericSeed) }
            }

            $isNum = $value -match '^\d'
            $presence = if ($isNum -and (RandInt 1 100) -le 92) { 'OK' } else { '' }
            $comment = if ($isNum) { '' } else { Pick $commentSeed }
            $doc = if ($isNum) { $null } else { @( $true, $false, $null )[(RandInt 0 2)] }
            $arrivalMarked = ((RandInt 1 100) -le 15)

            $dayValues[[string]$day] = $value
            $dayEntries[[string]$day] = [pscustomobject]@{
                Value = $value
                PresenceMark = $presence
                Comment = $comment
                DocumentAccepted = $doc
                ArrivalMarked = $arrivalMarked
            }
        }

        $months += [pscustomobject]@{
            MonthKey = $mk
            DayValues = [pscustomobject]$dayValues
            DayEntries = [pscustomobject]$dayEntries
        }
    }

    $timesheetRows += [pscustomobject]@{
        PersonId = $p.PersonId
        FullName = $p.FullName
        Specialty = $p.Specialty
        Rank = $p.Rank
        BrigadeName = $p.BrigadeName
        IsBrigadier = [bool]$p.IsBrigadier
        Months = $months
    }

    $primaryDate = $startDate.AddDays((RandInt 0 90))
    $brigadierName = if ($p.IsBrigadier) { $null } else { $leaders[$p.BrigadeName] }
    $n = RandInt 1 160

    $otRows += [pscustomobject]@{
        PersonId = $p.PersonId
        InstructionDate = $primaryDate.ToString('s')
        FullName = $p.FullName
        Specialty = $p.Specialty
        Rank = $p.Rank
        Profession = (Pick $professionSeed)
        InstructionType = $primaryType
        InstructionNumbers = "N-$n-25"
        RepeatPeriodMonths = 3
        IsBrigadier = [bool]$p.IsBrigadier
        BrigadierName = $brigadierName
        IsDismissed = ((RandInt 1 100) -le 4)
        IsPendingRepeat = $false
        IsRepeatCompleted = $false
    }

    $repDate = $primaryDate.AddMonths(3)
    $repIdx = 1
    while ($repDate -le $endDate) {
        $late = $repDate -ge [datetime]'2026-01-01'
        if ($late) {
            $pending = ((RandInt 1 100) -le 40)
            $completed = (-not $pending) -and ((RandInt 1 100) -le 70)
        } else {
            $completed = ((RandInt 1 100) -le 90)
            $pending = (-not $completed) -and ((RandInt 1 100) -le 60)
        }

        $otRows += [pscustomobject]@{
            PersonId = $p.PersonId
            InstructionDate = $repDate.ToString('s')
            FullName = $p.FullName
            Specialty = $p.Specialty
            Rank = $p.Rank
            Profession = (Pick $professionSeed)
            InstructionType = $repeatType
            InstructionNumbers = "R-$n-$repIdx"
            RepeatPeriodMonths = 3
            IsBrigadier = [bool]$p.IsBrigadier
            BrigadierName = $brigadierName
            IsDismissed = $false
            IsPendingRepeat = $pending
            IsRepeatCompleted = $completed
        }

        $repDate = $repDate.AddMonths(3)
        $repIdx++
    }
}
$co.TimesheetPeople = $timesheetRows
$co.OtJournal = $otRows | Sort-Object { [datetime]$_.InstructionDate }, FullName

$actionPool = @($co.ProductionJournal | Select-Object -ExpandProperty ActionName -Unique | Where-Object { $_ })
if ($actionPool.Count -eq 0) { $actionPool = @('WORK') }
$deviationPool = @($co.ProductionJournal | Select-Object -ExpandProperty Deviations -Unique | Where-Object { $_ })
if ($deviationPool.Count -eq 0) { $deviationPool = @('ok', 'axis +3 mm', 'axis +5 mm') }
$weatherPool = @($co.ProductionJournal | Select-Object -ExpandProperty Weather -Unique | Where-Object { $_ })
if ($weatherPool.Count -eq 0) { $weatherPool = @('clear') }
$actionPool = @($actionPool | ForEach-Object { FirstString $_ } | Where-Object { $_ } | Select-Object -Unique)
$deviationPool = @($deviationPool | ForEach-Object { FirstString $_ } | Where-Object { $_ } | Select-Object -Unique)
$weatherPool = @($weatherPool | ForEach-Object { FirstString $_ } | Where-Object { $_ } | Select-Object -Unique)
if ($actionPool.Count -eq 0) { $actionPool = @('WORK') }
if ($deviationPool.Count -eq 0) { $deviationPool = @('ok') }
if ($weatherPool.Count -eq 0) { $weatherPool = @('clear') }

function BuildWeather([datetime]$d, [object[]]$pool) {
    $template = Pick $pool
    $temp = switch ($d.Month) {
        { $_ -in 12, 1, 2 } { RandInt -10 2 }
        { $_ -in 3, 4 } { RandInt 0 12 }
        { $_ -in 5, 6 } { RandInt 12 24 }
        { $_ -in 7, 8 } { RandInt 20 32 }
        { $_ -in 9, 10 } { RandInt 8 20 }
        default { RandInt 0 10 }
    }
    if ($template -match '[-+]?\d+') {
        return ([regex]::Replace($template, '[-+]?\d+', "$temp", 1))
    }
    return "$temp C, $template"
}

$prodRows = @()
$d = $startDate
while ($d -le $endDate) {
    if ($d.DayOfWeek -notin @([System.DayOfWeek]::Saturday, [System.DayOfWeek]::Sunday)) {
        $perDay = RandInt 2 6
        for ($i = 0; $i -lt $perDay; $i++) {
            $g = Pick $mainGroups
            $parts = @()
            foreach ($m in (PickMany $extendedMaterials[$g] 1 2)) { $parts += "$m - $(RandInt 2 10)" }
            $blocks = @(PickMany @(1..$blocksCount) 1 2 | Sort-Object)
            $marks = PickMany $marksByGroup[$g] 1 2
            $remain = @()
            foreach ($b in 1..$blocksCount) { $remain += "B$b $(Pick $marksByGroup[$g]): left $(RandInt 0 28)" }

            $prodRows += [pscustomobject]@{
                Date = $d.ToString('s')
                ActionName = (Pick $actionPool)
                WorkName = $g
                ElementsText = ($parts -join '; ')
                BlocksText = ($blocks -join ', ')
                MarksText = ($marks -join ', ')
                BrigadeName = (Pick $brigadeSeed)
                Weather = (BuildWeather $d $weatherPool)
                Deviations = (Pick $deviationPool)
                RequiresHiddenWorkAct = ((RandInt 1 100) -le 35)
                RemainingInfo = ($remain -join '; ')
            }
        }
    }
    $d = $d.AddDays(1)
}
$co.ProductionJournal = $prodRows | Sort-Object { [datetime]$_.Date }

$inspectionSeed = @($co.InspectionJournal)
if ($inspectionSeed.Count -eq 0) {
    $inspectionSeed = @(
        [pscustomobject]@{
            JournalName = 'Inspection Log'
            InspectionName = 'Weekly check'
            ReminderPeriodDays = 7
            Notes = 'General control'
        }
    )
}
$inspectionRows = @()
for ($i = 0; $i -lt 120; $i++) {
    $seed = $inspectionSeed[$i % $inspectionSeed.Count]
    $rStart = $startDate.AddDays((RandInt 0 364))
    $period = if ($seed.ReminderPeriodDays -gt 0) { [int]$seed.ReminderPeriodDays } else { 7 }
    $lastDone = if ((RandInt 1 100) -le 85) { $rStart.AddDays((RandInt 0 364)) } else { $null }
    if ($lastDone -and $lastDone -gt $endDate) { $lastDone = $endDate.AddDays(- (RandInt 0 20)) }

    $inspectionRows += [pscustomobject]@{
        JournalName = $seed.JournalName
        InspectionName = "$($seed.InspectionName) #$($i + 1)"
        ReminderStartDate = $rStart.ToString('s')
        ReminderPeriodDays = $period
        LastCompletedDate = if ($lastDone) { $lastDone.ToString('s') } else { $null }
        Notes = if ($seed.Notes) { $seed.Notes } else { 'checklist' }
    }
}
$co.InspectionJournal = $inspectionRows | Sort-Object JournalName, InspectionName

$prefixByGroup = @{}
foreach ($g in ($mainGroups + $extraGroups)) {
    $prefix = ([regex]::Match($g, '^[^\d\s]+').Value)
    if (-not $prefix) { $prefix = 'MAT' }
    if ($prefix.Length -gt 3) { $prefix = $prefix.Substring(0,3) }
    $prefixByGroup[$g] = $prefix.ToUpperInvariant()
}

$journalRows = @()
$arrivalSheetName = [string]::Concat([char]0x041F, [char]0x0440, [char]0x0438, [char]0x0445, [char]0x043E, [char]0x0434)
$ttnCounter = 1
$d = $startDate
while ($d -le $endDate) {
    if ($d.DayOfWeek -notin @([System.DayOfWeek]::Saturday, [System.DayOfWeek]::Sunday)) {
        if ((RandInt 1 100) -le 95) {
            $arrivals = RandInt 2 6
            for ($a = 0; $a -lt $arrivals; $a++) {
                $isMain = ((RandInt 1 100) -le 74)
                if ($isMain) {
                    $g = Pick $mainGroups
                    $cat = $mainCategory
                    $sub = ''
                } else {
                    $g = Pick $extraGroups
                    $cat = $extraCategory
                    $sub = if ($extraGroups[0..1] -contains $g) { $extraSubA } else { $extraSubB }
                }

                $rowCount = if ($isMain) { RandInt 3 7 } else { RandInt 2 4 }
                $materials = PickMany $extendedMaterials[$g] $rowCount $rowCount
                $ttn = "{0}/{1}" -f $d.ToString('yyMMdd'), $ttnCounter
                $passport = "PS-$($d.ToString('yyMMdd'))-$ttnCounter"

                foreach ($m in $materials) {
                    $qty = if ($mainGroups -contains $g) { RandInt 2 24 } else { RandInt 5 650 }
                    $unit = $unitByGroup[$g]
                    $stb = if ($co.StbByGroup.PSObject.Properties[$g]) { $co.StbByGroup.$g } else { 'STD' }
                    $sup = if ($co.SupplierByGroup.PSObject.Properties[$g]) { $co.SupplierByGroup.$g } else { 'SUP' }
                    $vol = if ($unit -eq 'pc') {
                        [math]::Round($qty * (RandInt 20 140) / 100.0, 2)
                    } elseif ($unit -eq 'kg') {
                        [math]::Round($qty / 1000.0, 2)
                    } else {
                        [math]::Round($qty, 2)
                    }

                    $journalRows += [pscustomobject]@{
                        SheetName = 'Приход'
                        Date = $d.ToString('s')
                        ObjectName = FirstString $co.Name 'Object'
                        Category = FirstString $cat 'MAIN'
                        SubCategory = FirstString $sub ''
                        MaterialGroup = FirstString $g ''
                        MaterialName = FirstString $m ''
                        Unit = FirstString $unit 'pc'
                        Quantity = [double]$qty
                        Passport = $passport
                        Ttn = $ttn
                        Stb = FirstString $stb 'STD'
                        Supplier = FirstString $sup 'SUP'
                        Position = "{0}-{1}" -f $prefixByGroup[$g], (RandInt 1 99)
                        Volume = $vol.ToString('0.##', [System.Globalization.CultureInfo]::InvariantCulture)
                    }
                }
                $ttnCounter++
            }
        }
    }
    $d = $d.AddDays(1)
}
foreach ($row in $journalRows) {
    $row.SheetName = $arrivalSheetName
}
$state.Journal = $journalRows | Sort-Object { [datetime]$_.Date }, Ttn, MaterialGroup, MaterialName

$archiveGroups = [ordered]@{}
foreach ($r in $state.Journal) {
    if (-not $archiveGroups.Contains($r.MaterialGroup)) {
        $archiveGroups[$r.MaterialGroup] = New-Object 'System.Collections.Generic.HashSet[string]'
    }
    [void]$archiveGroups[$r.MaterialGroup].Add($r.MaterialName)
}
$archiveMaterials = [ordered]@{}
foreach ($g in $archiveGroups.Keys) { $archiveMaterials[$g] = @($archiveGroups[$g] | Sort-Object) }
$co.Archive = [pscustomobject]@{
    Groups = @($archiveGroups.Keys | Sort-Object)
    Materials = [pscustomobject]$archiveMaterials
    Units = @($state.Journal | Select-Object -ExpandProperty Unit -Unique | Where-Object { $_ } | Sort-Object)
    Suppliers = @($state.Journal | Select-Object -ExpandProperty Supplier -Unique | Where-Object { $_ } | Sort-Object)
    Passports = @($state.Journal | Select-Object -ExpandProperty Passport -Unique | Where-Object { $_ } | Sort-Object)
    Stb = @($state.Journal | Select-Object -ExpandProperty Stb -Unique | Where-Object { $_ } | Sort-Object)
}

$co.ArrivalHistory = @()
$co.PdfDocuments = @()
$co.EstimateDocuments = @()

$json = $state | ConvertTo-Json -Depth 100
[System.IO.File]::WriteAllText($path, $json, (New-Object System.Text.UTF8Encoding($false)))

Write-Output 'Large test data generated.'
