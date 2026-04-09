$ErrorActionPreference = 'Stop'

$outPath = Join-Path $PSScriptRoot 'data.json'
$rng = [System.Random]::new(20260409)

function RandInt([int]$min, [int]$max) {
    return $rng.Next($min, $max + 1)
}

function Pick([object[]]$items) {
    return $items[$rng.Next(0, $items.Count)]
}

function PickMany([object[]]$items, [int]$count) {
    return $items | Sort-Object { $rng.Next() } | Select-Object -First $count
}

function New-Dictionary {
    return [ordered]@{}
}

$types = [ordered]@{
    'Р РёРіРµР»Рё' = [ordered]@{
        Marks = @('+0.080', '+3.220', '+6.450')
        Materials = @('Р РћРџ 4.26-30', 'Р Рћ 23-4', 'Р Р”Рџ 4.56-50', 'Р РћРџ 4.35-3', 'Р Р“-01')
    }
    'РџР»РёС‚С‹ РїРµСЂРµРєСЂС‹С‚РёСЏ' = [ordered]@{
        Marks = @('0.000', '+3.000', '+6.000')
        Materials = @('РџРљ56.15-12', 'РџРљ56.12-10', 'РџРљ56.15-10', 'РџРљ36.15-8', 'РџР‘-01')
    }
    'Р”РёР°С„СЂР°РіРјС‹' = [ordered]@{
        Marks = @('0.000', '+3.000', '+6.000')
        Materials = @('Р”Р¤-01', 'Р”Р¤-02', 'Р”Р¤-03', 'Р”Р¤-04', 'Р”Р¤-05')
    }
    'РљРѕР»РѕРЅРЅС‹' = [ordered]@{
        Marks = @('0.000', '+3.300', '+6.600')
        Materials = @('Рљ-01', 'Рљ-02', 'Рљ-03', 'Рљ-04', 'Рљ-05')
    }
    'Р›РµСЃС‚РЅРёС‡РЅС‹Рµ РјР°СЂС€Рё' = [ordered]@{
        Marks = @('0.000', '+3.000', '+6.000')
        Materials = @('Р›Рњ-01', 'Р›Рњ-02', 'Р›Рњ-03', 'Р›Рњ-04', 'Р›Рњ-05')
    }
}

$blocks = @(1, 2, 3)
$arrivalMonths = @(
    [datetime]'2025-11-01',
    [datetime]'2025-12-01',
    [datetime]'2026-01-01',
    [datetime]'2026-02-01',
    [datetime]'2026-03-01'
)

$project = [ordered]@{
    Demand = New-Dictionary
    Archive = [ordered]@{
        Groups = @()
        Materials = New-Dictionary
        Units = @()
        Suppliers = @()
        Passports = @()
        Stb = @()
    }
    Name = 'РЎРѕРєРѕР»'
    BlocksCount = 3
    HasBasement = $true
    SameFloorsInBlocks = $true
    FloorsPerBlock = 4
    FloorsByBlock = New-Dictionary
    BlockAxesByNumber = [ordered]@{
        '1' = '1-11/Р›-Р“Р“'
        '2' = '12-22/Р›-Р“Р“'
        '3' = '23-33/Р›-Р“Р“'
    }
    FullObjectName = 'РЎС‚СЂРѕРёС‚РµР»СЊСЃС‚РІРѕ С€РєРѕР»С‹ "РЎРѕРєРѕР»"'
    GeneralContractorRepresentative = 'РџРµС‚СЂРѕРІ РРіРѕСЂСЊ Р’РёРєС‚РѕСЂРѕРІРёС‡'
    TechnicalSupervisorRepresentative = 'РљР»РёРјРѕРІ РЎРµСЂРіРµР№ РџР°РІР»РѕРІРёС‡'
    ProjectOrganizationRepresentative = 'РРІР°РЅРѕРІ РђР»РµРєСЃРµР№ Р РѕРјР°РЅРѕРІРёС‡'
    ProjectDocumentationName = 'Р Р°Р·РґРµР» РљР–. Р Р°Р±РѕС‡Р°СЏ РґРѕРєСѓРјРµРЅС‚Р°С†РёСЏ'
    MasterNames = @('РЎРёРґРѕСЂРѕРІ РџР°РІРµР»', 'РќРёРєРёС‚РёРЅ РђСЂС‚РµРј')
    ForemanNames = @('РљСѓР·РЅРµС†РѕРІ РРІР°РЅ', 'РњРµР»СЊРЅРёРє Р”РјРёС‚СЂРёР№')
    ResponsibleForeman = 'РљСѓР·РЅРµС†РѕРІ РРІР°РЅ'
    SiteManagerName = 'РћСЂР»РѕРІ РњР°РєСЃРёРј'
    MaterialNamesByGroup = New-Dictionary
    StbByGroup = New-Dictionary
    SupplierByGroup = New-Dictionary
    MaterialGroups = @()
    MaterialCatalog = @()
    MaterialTreeSplitRules = New-Dictionary
    AutoSplitMaterialNames = @('РџРљ56.15-12', 'Р РћРџ 4.26-30', 'Р Р”Рџ 4.56-50')
    ArrivalHistory = @()
    SummaryVisibleGroups = @()
    SummaryMarksByGroup = New-Dictionary
    OtJournal = @()
    TimesheetPeople = @()
    ProductionJournal = @()
    ProductionAutoFillSettings = [ordered]@{
        MinQuantityPerRow = 4
        MaxQuantityPerRow = 8
        MinRowsPerRun = 4
        TargetRowsPerRun = 5
        MaxRowsPerRun = 6
        MaxItemsPerRow = 2
        PreferSelectedTypeOnly = $true
        UseBalancedDistribution = $true
        PreferDemandDeficit = $true
        RespectSelectedBlocksAndMarks = $true
        AllowMixedMaterialsInRow = $true
    }
    InspectionJournal = @()
    PdfDocuments = @()
    EstimateDocuments = @()
    UiSettings = [ordered]@{
        DisableTree = $false
        PinTreeByDefault = $false
        ShowReminderPopup = $true
        ReminderSnoozeMinutes = 15
        HideReminderDetails = $false
    }
}

$totalNeedByMaterial = New-Dictionary

foreach ($typeName in $types.Keys) {
    $meta = $types[$typeName]
    $marks = @($meta.Marks)
    $materials = @($meta.Materials)

    $project.MaterialNamesByGroup[$typeName] = $materials
    $project.MaterialGroups += [ordered]@{ Name = $typeName; Items = $materials }
    $project.StbByGroup[$typeName] = 'РЎРўР‘ 1300'
    $project.SupplierByGroup[$typeName] = 'РћРђРћ РњР–Р‘'
    $project.SummaryVisibleGroups += $typeName
    $project.SummaryMarksByGroup[$typeName] = $marks

    foreach ($materialName in $materials) {
        $project.MaterialCatalog += [ordered]@{
            CategoryName = 'РћСЃРЅРѕРІРЅС‹Рµ'
            TypeName = $typeName
            SubTypeName = ''
            ExtraLevels = @()
            LevelMarks = $marks
            MaterialName = $materialName
        }

        $levels = New-Dictionary
        $mounted = New-Dictionary
        $totalNeed = 0

        foreach ($block in $blocks) {
            $levelRow = New-Dictionary
            $mountedRow = New-Dictionary
            foreach ($mark in $marks) {
                $need = RandInt 8 20
                $installed = RandInt 0 ($need - 2)
                $levelRow[$mark] = [double]$need
                $mountedRow[$mark] = [double]$installed
                $totalNeed += $need
            }
            $levels["$block"] = $levelRow
            $mounted["$block"] = $mountedRow
        }

        $demandKey = "$typeName::$materialName"
        $project.Demand[$demandKey] = [ordered]@{
            Unit = 'С€С‚'
            Levels = $levels
            MountedLevels = $mounted
            Floors = New-Dictionary
            MountedFloors = New-Dictionary
        }
        $totalNeedByMaterial[$demandKey] = $totalNeed
    }
}

$journal = @()
$ttnCounter = 1
foreach ($demandKey in $totalNeedByMaterial.Keys) {
    $parts = $demandKey.Split('::')
    $typeName = $parts[0]
    $materialName = $parts[1]
    $needTotal = [int]$totalNeedByMaterial[$demandKey]

    $isOverage = $rng.NextDouble() -lt 0.35
    $factor = if ($isOverage) { 1.10 + ($rng.NextDouble() * 0.25) } else { 0.80 + ($rng.NextDouble() * 0.22) }
    $targetArrival = [Math]::Max(6, [int][Math]::Round($needTotal * $factor))
    $remaining = $targetArrival

    for ($monthIndex = 0; $monthIndex -lt $arrivalMonths.Count; $monthIndex++) {
        $monthStart = $arrivalMonths[$monthIndex]
        $monthDate = $monthStart.AddDays((RandInt 1 24))
        if ($monthIndex -eq $arrivalMonths.Count - 1) {
            $qty = $remaining
        }
        else {
            $basePart = [Math]::Round($targetArrival / $arrivalMonths.Count)
            $qty = [Math]::Max(0, [int]($basePart + (RandInt -2 3)))
            $remaining -= $qty
        }

        if ($qty -le 0) { continue }

        $journal += [ordered]@{
            SheetName = 'РџСЂРёС…РѕРґ'
            Date = $monthDate.ToString('s')
            ObjectName = $project.Name
            Category = 'РћСЃРЅРѕРІРЅС‹Рµ'
            SubCategory = ''
            MaterialGroup = $typeName
            MaterialName = $materialName
            Unit = 'С€С‚'
            Quantity = [double]$qty
            Passport = "РџРЎ-$($monthDate.ToString('yyMMdd'))-$ttnCounter"
            Ttn = "{0:yyMM}-{1:000}" -f $monthDate, $ttnCounter
            Stb = 'РЎРўР‘ 1300'
            Supplier = 'РћРђРћ РњР–Р‘'
            Position = "Рџ-$ttnCounter"
            Volume = [Math]::Round($qty * 0.1, 2).ToString([System.Globalization.CultureInfo]::InvariantCulture)
        }
        $ttnCounter++
    }
}

$people = @(
    'РРІР°РЅРѕРІ РЎРµСЂРіРµР№ Р’РёРєС‚РѕСЂРѕРІРёС‡', 'РџРµС‚СЂРѕРІ РџР°РІРµР» РђРЅРґСЂРµРµРІРёС‡', 'РЎРёРґРѕСЂРѕРІ РђР»РµРєСЃРµР№ РРіРѕСЂРµРІРёС‡', 'РЎРјРёСЂРЅРѕРІ РћР»РµРі РќРёРєРѕР»Р°РµРІРёС‡',
    'РљСѓР·РЅРµС†РѕРІ РР»СЊСЏ РњР°РєСЃРёРјРѕРІРёС‡', 'РћСЂР»РѕРІ РђРЅС‚РѕРЅ РЎРµСЂРіРµРµРІРёС‡', 'РњРµР»СЊРЅРёРє Р”РјРёС‚СЂРёР№ РџР°РІР»РѕРІРёС‡', 'Р СѓРґРµРЅРєРѕ РђСЂС‚РµРј РР»СЊРёС‡',
    'Р‘РµР»С‹Р№ РљРёСЂРёР»Р» Р РѕРјР°РЅРѕРІРёС‡', 'Р•РіРѕСЂРѕРІ Р’Р»Р°РґРёСЃР»Р°РІ РЎРµСЂРіРµРµРІРёС‡', 'Р–СѓРєРѕРІ РќРёРєРёС‚Р° РђР»РµРєСЃРµРµРІРёС‡', 'Р’РѕР»РєРѕРІ Р РѕРјР°РЅ РРіРѕСЂРµРІРёС‡',
    'Р¤РµРґРѕСЂРѕРІ Р”Р°РЅРёРёР» РџР°РІР»РѕРІРёС‡', 'РўРёС…РѕРЅРѕРІ РђР»РµРєСЃРµР№ РђСЂС‚РµРјРѕРІРёС‡', 'Р“РѕСЂР±СѓРЅРѕРІ Р•РіРѕСЂ РЎРµСЂРіРµРµРІРёС‡', 'РљР»РёРјРѕРІ РњР°РєСЃРёРј РџР°РІР»РѕРІРёС‡',
    'Р РѕРјР°РЅРѕРІ РђСЂС‚СѓСЂ РРіРѕСЂРµРІРёС‡', 'Р—Р°Р№С†РµРІ Р”РµРЅРёСЃ РђРЅРґСЂРµРµРІРёС‡', 'РЎРѕР»РѕРІСЊРµРІ РРіРѕСЂСЊ Р РѕРјР°РЅРѕРІРёС‡', 'РџРѕР»СЏРєРѕРІ Р”РјРёС‚СЂРёР№ РЎРµСЂРіРµРµРІРёС‡'
)
$specialties = @('РњРѕРЅС‚Р°Р¶РЅРёРє Р–Р‘Рљ', 'РђСЂРјР°С‚СѓСЂС‰РёРє', 'Р‘РµС‚РѕРЅС‰РёРє', 'Р­Р»РµРєС‚СЂРѕРјРѕРЅС‚Р°Р¶РЅРёРє', 'РЎРІР°СЂС‰РёРє')
$ranks = @('3', '4', '5', '6')
$brigades = @('Р‘СЂРёРіР°РґР° 1', 'Р‘СЂРёРіР°РґР° 2', 'Р‘СЂРёРіР°РґР° 3', 'Р‘СЂРёРіР°РґР° 4')

$today = Get-Date
$currentMonth = Get-Date -Year $today.Year -Month $today.Month -Day 1
$tsMonths = @($currentMonth.AddMonths(-2), $currentMonth.AddMonths(-1), $currentMonth, $currentMonth.AddMonths(1))

for ($i = 0; $i -lt $people.Count; $i++) {
    $personId = [guid]::NewGuid().ToString()
    $dailyHours = if (($i % 5) -eq 0) { 12 } else { 8 }
    $brigade = $brigades[$i % $brigades.Count]
    $isBrigadier = ($i % 5) -eq 0
    $specialty = Pick $specialties
    $rank = Pick $ranks

    $months = @()
    foreach ($month in $tsMonths) {
        $monthKey = $month.ToString('yyyy-MM')
        $dayValues = New-Dictionary
        $dayEntries = New-Dictionary
        $daysInMonth = [datetime]::DaysInMonth($month.Year, $month.Month)
        $isFutureMonth = ($month -eq $currentMonth.AddMonths(1))
        if (-not $isFutureMonth) {
            foreach ($day in 1..$daysInMonth) {
                $date = [datetime]::new($month.Year, $month.Month, $day)
                $isWeekend = $date.DayOfWeek -in @([System.DayOfWeek]::Saturday, [System.DayOfWeek]::Sunday)
                if ($isWeekend) {
                    $value = 'Р’'
                    $comment = ''
                    $docAccepted = $null
                }
                else {
                    $roll = RandInt 1 100
                    if ($roll -le 82) {
                        $value = "$dailyHours"
                        $comment = ''
                        $docAccepted = $null
                    }
                    elseif ($roll -le 90) {
                        $value = 'Рќ'
                        $comment = 'РћС‚СЃСѓС‚СЃС‚РІРёРµ Р±РµР· РїРѕРґС‚РІРµСЂР¶РґРµРЅРёСЏ'
                        $docAccepted = $false
                    }
                    else {
                        $value = 'Р‘'
                        $comment = 'Р‘РѕР»СЊРЅРёС‡РЅС‹Р№ Р»РёСЃС‚'
                        $docAccepted = $true
                    }
                }

                $dayValues["$day"] = $value
                $dayEntries["$day"] = [ordered]@{
                    Value = $value
                    PresenceMark = if ($value -match '^\d+$') { 'вњ”' } else { '' }
                    Comment = $comment
                    DocumentAccepted = $docAccepted
                    ArrivalMarked = $false
                }
            }
        }

        $months += [ordered]@{
            MonthKey = $monthKey
            DayValues = $dayValues
            DayEntries = $dayEntries
        }
    }

    $project.TimesheetPeople += [ordered]@{
        PersonId = $personId
        FullName = $people[$i]
        Specialty = $specialty
        Rank = $rank
        BrigadeName = $brigade
        IsBrigadier = $isBrigadier
        DailyWorkHours = $dailyHours
        Months = $months
    }

    $instructionDate = [datetime]'2025-11-15'.AddDays((RandInt 0 75))
    $project.OtJournal += [ordered]@{
        PersonId = $personId
        InstructionDate = $instructionDate.ToString('s')
        FullName = $people[$i]
        Specialty = $specialty
        Rank = $rank
        Profession = $specialty
        InstructionType = 'РџРµСЂРІРёС‡РЅС‹Р№ РЅР° СЂР°Р±РѕС‡РµРј РјРµСЃС‚Рµ'
        InstructionNumbers = "РџР -$($i + 101)"
        RepeatPeriodMonths = 3
        IsBrigadier = $isBrigadier
        BrigadierName = if ($isBrigadier) { '' } else { $brigade }
        IsDismissed = $false
        IsPendingRepeat = $false
        IsRepeatCompleted = $false
        IsScheduledRepeat = $false
    }

    if ($i -lt 6) {
        $project.OtJournal += [ordered]@{
            PersonId = $personId
            InstructionDate = (Get-Date).AddDays(-(RandInt 1 14)).ToString('s')
            FullName = $people[$i]
            Specialty = $specialty
            Rank = $rank
            Profession = $specialty
            InstructionType = 'РџРѕРІС‚РѕСЂРЅС‹Р№'
            InstructionNumbers = "РџР’-$($i + 201)"
            RepeatPeriodMonths = 3
            IsBrigadier = $isBrigadier
            BrigadierName = if ($isBrigadier) { '' } else { $brigade }
            IsDismissed = $false
            IsPendingRepeat = $true
            IsRepeatCompleted = $false
            IsScheduledRepeat = $false
        }
    }
}

$weatherByMonth = @{
    11 = '+3В°C, РїР°СЃРјСѓСЂРЅРѕ'
    12 = '-5В°C, СЃРЅРµРі'
    1  = '-8В°C, СЃРЅРµРі'
    2  = '-2В°C, СЏСЃРЅРѕ'
    3  = '+4В°C, РѕР±Р»Р°С‡РЅРѕ'
}
$deviations = @('РћС‚РєР»РѕРЅРµРЅРёР№ РЅРµС‚', 'РћС‚РєР»РѕРЅРµРЅРёРµ РѕС‚ СЂР°Р·Р±РёРІРѕС‡РЅС‹С… РѕСЃРµР№ +3 РјРј', 'РћС‚РєР»РѕРЅРµРЅРёРµ РѕС‚ СЂР°Р·Р±РёРІРѕС‡РЅС‹С… РѕСЃРµР№ +5 РјРј')

foreach ($monthStart in $arrivalMonths) {
    foreach ($typeName in $types.Keys) {
        $marks = @($types[$typeName].Marks)
        $materials = @($types[$typeName].Materials)
        $selected = PickMany $materials 2
        $elements = @()
        foreach ($material in $selected) {
            $elements += "$material - $(RandInt 3 9)"
        }
        $mark = Pick $marks
        $block = RandInt 1 3
        $remainingRows = @()
        foreach ($material in $selected) {
            $remainingRows += "$material: Р‘$block $mark вЂ” РѕСЃС‚Р°С‚РѕРє $(RandInt 0 30)"
        }

        $project.ProductionJournal += [ordered]@{
            Date = $monthStart.AddDays((RandInt 2 24)).ToString('s')
            ActionName = Pick @('РњРѕРЅС‚Р°Р¶', 'РљР»Р°РґРєР°', 'РЈСЃС‚СЂРѕР№СЃС‚РІРѕ')
            WorkName = $typeName
            ElementsText = ($elements -join '; ')
            BlocksText = "$block"
            MarksText = $mark
            BrigadeName = Pick $brigades
            Weather = $weatherByMonth[$monthStart.Month]
            Deviations = Pick $deviations
            RequiresHiddenWorkAct = ((RandInt 1 100) -le 30)
            RemainingInfo = ($remainingRows -join '; ')
            SuppressDateDisplay = $false
            SuppressWeatherDisplay = $false
        }
    }
}

$inspectionJournals = @(
    'Р–СѓСЂРЅР°Р» РѕСЃРјРѕС‚СЂР° Р»РµСЃРѕРІ',
    'Р–СѓСЂРЅР°Р» РѕСЃРјРѕС‚СЂР° РІСЂРµРјРµРЅРЅС‹С… РѕРіСЂР°Р¶РґРµРЅРёР№',
    'Р–СѓСЂРЅР°Р» РѕСЃРјРѕС‚СЂР° РіСЂСѓР·РѕР·Р°С…РІР°С‚РЅС‹С… РїСЂРёСЃРїРѕСЃРѕР±Р»РµРЅРёР№',
    'Р–СѓСЂРЅР°Р» РѕСЃРјРѕС‚СЂР° СЌР»РµРєС‚СЂРѕРёРЅСЃС‚СЂСѓРјРµРЅС‚Р°',
    'Р–СѓСЂРЅР°Р» РѕСЃРјРѕС‚СЂР° РїРѕРґРјРѕСЃС‚РµР№'
)
$inspectionNames = @('Р•Р¶РµРЅРµРґРµР»СЊРЅС‹Р№ РѕСЃРјРѕС‚СЂ', 'РџР»Р°РЅРѕРІС‹Р№ РѕСЃРјРѕС‚СЂ РїРµСЂРµРґ СЃРјРµРЅРѕР№', 'РљРѕРЅС‚СЂРѕР»СЊ СЃРѕСЃС‚РѕСЏРЅРёСЏ', 'Р•Р¶РµРјРµСЃСЏС‡РЅС‹Р№ РѕСЃРјРѕС‚СЂ')

for ($i = 0; $i -lt 10; $i++) {
    $start = [datetime]'2026-01-10'.AddDays((RandInt 0 80))
    $period = Pick @(7, 10, 14, 30)
    $lastDone = if ($i % 3 -eq 0) { (Get-Date).AddDays(-(RandInt 3 35)) } else { $null }
    $project.InspectionJournal += [ordered]@{
        JournalName = $inspectionJournals[$i % $inspectionJournals.Count]
        InspectionName = "$($inspectionNames[$i % $inspectionNames.Count]) в„–$($i + 1)"
        ReminderStartDate = $start.ToString('s')
        ReminderPeriodDays = $period
        LastCompletedDate = if ($lastDone) { $lastDone.ToString('s') } else { $null }
        Notes = Pick @('РџСЂРѕРІРµСЂРµРЅРѕ РїРѕ С‡РµРє-Р»РёСЃС‚Сѓ', 'Р‘РµР· Р·Р°РјРµС‡Р°РЅРёР№', 'РўСЂРµР±СѓРµС‚СЃСЏ РїРѕРґРїРёСЃСЊ РјР°СЃС‚РµСЂР°', 'Р—Р°РјРµС‡Р°РЅРёСЏ СѓСЃС‚СЂР°РЅРёС‚СЊ РґРѕ РєРѕРЅС†Р° СЃРјРµРЅС‹')
        IsCompletionHistory = ($i -ge 7)
    }
}

$archiveGroups = $journal | Select-Object -ExpandProperty MaterialGroup -Unique | Sort-Object
$archiveMaterials = New-Dictionary
foreach ($groupName in $archiveGroups) {
    $archiveMaterials[$groupName] = @(
        $journal | Where-Object { $_.MaterialGroup -eq $groupName } | Select-Object -ExpandProperty MaterialName -Unique | Sort-Object
    )
}

$project.Archive = [ordered]@{
    Groups = @($archiveGroups)
    Materials = $archiveMaterials
    Units = @($journal | Select-Object -ExpandProperty Unit -Unique | Sort-Object)
    Suppliers = @($journal | Select-Object -ExpandProperty Supplier -Unique | Sort-Object)
    Passports = @($journal | Select-Object -ExpandProperty Passport -Unique | Sort-Object)
    Stb = @($journal | Select-Object -ExpandProperty Stb -Unique | Sort-Object)
}

$state = [ordered]@{
    CurrentObject = $project
    Journal = $journal
}

$json = $state | ConvertTo-Json -Depth 100
[System.IO.File]::WriteAllText($outPath, $json, [System.Text.UTF8Encoding]::new($false))
Write-Output "Р“РѕС‚РѕРІРѕ: С‚РµСЃС‚РѕРІС‹Рµ РґР°РЅРЅС‹Рµ Р·Р°РїРёСЃР°РЅС‹ РІ $outPath"
