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

function U([int[]]$codes) {
    return -join ($codes | ForEach-Object { [char]$_ })
}

$strCategoryMain = U @(0x041E,0x0441,0x043D,0x043E,0x0432,0x043D,0x044B,0x0435) # Основные
$strUnitPieces = U @(0x0448,0x0442) # шт
$strRepeat = U @(0x041F,0x043E,0x0432,0x0442,0x043E,0x0440,0x043D,0x044B,0x0439) # Повторный

$types = [ordered]@{
    'Rigeli' = [ordered]@{
        Marks = @('+0.080', '+3.220', '+6.450')
        Materials = @('ROP-4.26-30', 'RO-23-4', 'RDP-4.56-50', 'ROP-4.35-3', 'RG-01')
    }
    'Plity' = [ordered]@{
        Marks = @('0.000', '+3.000', '+6.000')
        Materials = @('PK56.15-12', 'PK56.12-10', 'PK56.15-10', 'PK36.15-8', 'PB-01')
    }
    'Diafragmy' = [ordered]@{
        Marks = @('0.000', '+3.000', '+6.000')
        Materials = @('DF-01', 'DF-02', 'DF-03', 'DF-04', 'DF-05')
    }
    'Kolonny' = [ordered]@{
        Marks = @('0.000', '+3.300', '+6.600')
        Materials = @('K-01', 'K-02', 'K-03', 'K-04', 'K-05')
    }
    'Marshi' = [ordered]@{
        Marks = @('0.000', '+3.000', '+6.000')
        Materials = @('LM-01', 'LM-02', 'LM-03', 'LM-04', 'LM-05')
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
    Name = 'Sokol'
    BlocksCount = 3
    HasBasement = $true
    SameFloorsInBlocks = $true
    FloorsPerBlock = 4
    FloorsByBlock = New-Dictionary
    BlockAxesByNumber = [ordered]@{
        '1' = '1-11/L-GG'
        '2' = '12-22/L-GG'
        '3' = '23-33/L-GG'
    }
    FullObjectName = 'Test object Sokol School'
    GeneralContractorRepresentative = 'Petrov I.V.'
    TechnicalSupervisorRepresentative = 'Klimov S.P.'
    ProjectOrganizationRepresentative = 'Ivanov A.R.'
    ProjectDocumentationName = 'KJ working docs'
    MasterNames = @('Sidorov P.', 'Nikitin A.')
    ForemanNames = @('Kuznetsov I.', 'Melnik D.')
    ResponsibleForeman = 'Kuznetsov I.'
    SiteManagerName = 'Orlov M.'
    MaterialNamesByGroup = New-Dictionary
    StbByGroup = New-Dictionary
    SupplierByGroup = New-Dictionary
    MaterialGroups = @()
    MaterialCatalog = @()
    MaterialTreeSplitRules = New-Dictionary
    AutoSplitMaterialNames = @('PK56.15-12', 'ROP-4.26-30', 'RDP-4.56-50')
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
    $project.StbByGroup[$typeName] = 'STB-1300'
    $project.SupplierByGroup[$typeName] = 'MJB'
    $project.SummaryVisibleGroups += $typeName
    $project.SummaryMarksByGroup[$typeName] = $marks

    foreach ($materialName in $materials) {
        $project.MaterialCatalog += [ordered]@{
            CategoryName = $strCategoryMain
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
            Unit = $strUnitPieces
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
    $factor = if ($rng.NextDouble() -lt 0.35) { 1.10 + ($rng.NextDouble() * 0.25) } else { 0.80 + ($rng.NextDouble() * 0.22) }
    $targetArrival = [Math]::Max(6, [int][Math]::Round($needTotal * $factor))
    $remaining = $targetArrival
    for ($monthIndex = 0; $monthIndex -lt $arrivalMonths.Count; $monthIndex++) {
        $monthDate = $arrivalMonths[$monthIndex].AddDays((RandInt 1 24))
        if ($monthIndex -eq $arrivalMonths.Count - 1) { $qty = $remaining } else { $qty = [Math]::Max(0, [int]([Math]::Round($targetArrival / $arrivalMonths.Count) + (RandInt -2 3))); $remaining -= $qty }
        if ($qty -le 0) { continue }
        $journal += [ordered]@{
            SheetName = 'Prikhod'
            Date = $monthDate.ToString('s')
            ObjectName = $project.Name
            Category = $strCategoryMain
            SubCategory = ''
            MaterialGroup = $typeName
            MaterialName = $materialName
            Unit = $strUnitPieces
            Quantity = [double]$qty
            Passport = "PS-$($monthDate.ToString('yyMMdd'))-$ttnCounter"
            Ttn = "{0:yyMM}-{1:000}" -f $monthDate, $ttnCounter
            Stb = 'STB-1300'
            Supplier = 'MJB'
            Position = "P-$ttnCounter"
            Volume = [Math]::Round($qty * 0.1, 2).ToString([System.Globalization.CultureInfo]::InvariantCulture)
        }
        $ttnCounter++
    }
}

$people = @(
    'Ivanov S.V.', 'Petrov P.A.', 'Sidorov A.I.', 'Smirnov O.N.', 'Kuznetsov I.M.',
    'Orlov A.S.', 'Melnik D.P.', 'Rudenko A.I.', 'Bely K.R.', 'Egorov V.S.',
    'Zhukov N.A.', 'Volkov R.I.', 'Fedorov D.P.', 'Tikhonov A.A.', 'Gorbunov E.S.',
    'Klimov M.P.', 'Romanov A.I.', 'Zaitsev D.A.', 'Solovev I.R.', 'Polyakov D.S.'
)
$specialties = @('Montazhnik', 'Armaturshik', 'Betonshik', 'Elektromontazhnik', 'Svarshik')
$ranks = @('3', '4', '5', '6')
$brigades = @('Brigada 1', 'Brigada 2', 'Brigada 3', 'Brigada 4')

$today = Get-Date
$currentMonth = Get-Date -Year $today.Year -Month $today.Month -Day 1
$tsMonths = @($currentMonth.AddMonths(-2), $currentMonth.AddMonths(-1), $currentMonth, $currentMonth.AddMonths(1))

for ($i = 0; $i -lt $people.Count; $i++) {
    $personId = [guid]::NewGuid().ToString()
    $dailyHours = if (($i % 5) -eq 0) { 12 } else { 8 }
    $brigade = $brigades[$i % $brigades.Count]
    $isBrigadier = ($i % 5) -eq 0
    $specialty = Pick $specialties

    $months = @()
    foreach ($month in $tsMonths) {
        $monthKey = $month.ToString('yyyy-MM')
        $dayValues = New-Dictionary
        $dayEntries = New-Dictionary
        if ($month -ne $currentMonth.AddMonths(1)) {
            $daysInMonth = [datetime]::DaysInMonth($month.Year, $month.Month)
            foreach ($day in 1..$daysInMonth) {
                $d = [datetime]::new($month.Year, $month.Month, $day)
                $isWeekend = $d.DayOfWeek -in @([System.DayOfWeek]::Saturday, [System.DayOfWeek]::Sunday)
                if ($isWeekend) { $value = 'V'; $doc = $null; $comment = '' }
                else {
                    $roll = RandInt 1 100
                    if ($roll -le 82) { $value = "$dailyHours"; $doc = $null; $comment = '' }
                    elseif ($roll -le 90) { $value = 'N'; $doc = $false; $comment = 'No document yet' }
                    else { $value = 'B'; $doc = $true; $comment = 'Sick leave' }
                }
                $dayValues["$day"] = $value
                $dayEntries["$day"] = [ordered]@{ Value = $value; PresenceMark = if ($value -match '^\d+$') { 'ok' } else { '' }; Comment = $comment; DocumentAccepted = $doc; ArrivalMarked = $false }
            }
        }
        $months += [ordered]@{ MonthKey = $monthKey; DayValues = $dayValues; DayEntries = $dayEntries }
    }

    $project.TimesheetPeople += [ordered]@{
        PersonId = $personId
        FullName = $people[$i]
        Specialty = $specialty
        Rank = Pick $ranks
        BrigadeName = $brigade
        IsBrigadier = $isBrigadier
        DailyWorkHours = $dailyHours
        Months = $months
    }

    $project.OtJournal += [ordered]@{
        PersonId = $personId
        InstructionDate = ([datetime]'2025-11-15').AddDays((RandInt 0 75)).ToString('s')
        FullName = $people[$i]
        Specialty = $specialty
        Rank = Pick $ranks
        Profession = $specialty
        InstructionNumbers = "PR-$($i + 101)"
        RepeatPeriodMonths = 3
        IsBrigadier = $isBrigadier
        BrigadierName = if ($isBrigadier) { '' } else { $brigade }
        IsDismissed = $false
        IsPendingRepeat = $false
        IsRepeatCompleted = $false
        IsScheduledRepeat = $false
    }
}

for ($i = 0; $i -lt 6; $i++) {
    $project.OtJournal += [ordered]@{
        PersonId = $project.TimesheetPeople[$i].PersonId
        InstructionDate = (Get-Date).AddDays(-(RandInt 1 14)).ToString('s')
        FullName = $project.TimesheetPeople[$i].FullName
        Specialty = $project.TimesheetPeople[$i].Specialty
        Rank = $project.TimesheetPeople[$i].Rank
        Profession = $project.TimesheetPeople[$i].Specialty
        InstructionType = $strRepeat
        InstructionNumbers = "PV-$($i + 201)"
        RepeatPeriodMonths = 3
        IsBrigadier = $project.TimesheetPeople[$i].IsBrigadier
        BrigadierName = $project.TimesheetPeople[$i].BrigadeName
        IsDismissed = $false
        IsPendingRepeat = $true
        IsRepeatCompleted = $false
        IsScheduledRepeat = $false
    }
}

$weatherByMonth = @{ 11 = '+3C cloudy'; 12 = '-5C snow'; 1 = '-8C snow'; 2 = '-2C clear'; 3 = '+4C cloudy' }
$dev = @('No deviations', 'Axis shift +3 mm', 'Axis shift +5 mm')

foreach ($monthStart in $arrivalMonths) {
    foreach ($typeName in $types.Keys) {
        $selected = PickMany @($types[$typeName].Materials) 2
        $mark = Pick @($types[$typeName].Marks)
        $block = RandInt 1 3
        $elements = @($selected | ForEach-Object { "$_ - $(RandInt 3 9)" }) -join '; '
        $remaining = @($selected | ForEach-Object { "$($_): B$block $mark remain $(RandInt 0 30)" }) -join '; '
        $project.ProductionJournal += [ordered]@{
            Date = $monthStart.AddDays((RandInt 2 24)).ToString('s')
            ActionName = Pick @('Montazh', 'Kladka', 'Ustroistvo')
            WorkName = $typeName
            ElementsText = $elements
            BlocksText = "$block"
            MarksText = $mark
            BrigadeName = Pick $brigades
            Weather = $weatherByMonth[$monthStart.Month]
            Deviations = Pick $dev
            RequiresHiddenWorkAct = ((RandInt 1 100) -le 30)
            RemainingInfo = $remaining
            SuppressDateDisplay = $false
            SuppressWeatherDisplay = $false
        }
    }
}

$inspectionJournals = @('Scaffold', 'Fences', 'Lifting', 'PowerTools', 'Platforms')
$inspectionNames = @('Weekly check', 'Shift check', 'Status control', 'Monthly check')
for ($i = 0; $i -lt 10; $i++) {
    $start = ([datetime]'2026-01-10').AddDays((RandInt 0 80))
    $period = Pick @(7, 10, 14, 30)
    $lastDone = if ($i % 3 -eq 0) { (Get-Date).AddDays(-(RandInt 3 35)) } else { $null }
    $project.InspectionJournal += [ordered]@{
        JournalName = $inspectionJournals[$i % $inspectionJournals.Count]
        InspectionName = "$($inspectionNames[$i % $inspectionNames.Count]) #$($i + 1)"
        ReminderStartDate = $start.ToString('s')
        ReminderPeriodDays = $period
        LastCompletedDate = if ($lastDone) { $lastDone.ToString('s') } else { $null }
        Notes = Pick @('Checklist complete', 'No remarks', 'Need master sign', 'Fix before end of shift')
        IsCompletionHistory = ($i -ge 7)
    }
}

$archiveGroups = @($journal | ForEach-Object { $_['MaterialGroup'] } | Sort-Object -Unique)
$archiveMaterials = New-Dictionary
foreach ($g in $archiveGroups) {
    $archiveMaterials[$g] = @(
        $journal |
            Where-Object { $_['MaterialGroup'] -eq $g } |
            ForEach-Object { $_['MaterialName'] } |
            Sort-Object -Unique
    )
}
$project.Archive = [ordered]@{
    Groups = @($archiveGroups)
    Materials = $archiveMaterials
    Units = @($journal | ForEach-Object { $_['Unit'] } | Sort-Object -Unique)
    Suppliers = @($journal | ForEach-Object { $_['Supplier'] } | Sort-Object -Unique)
    Passports = @($journal | ForEach-Object { $_['Passport'] } | Sort-Object -Unique)
    Stb = @($journal | ForEach-Object { $_['Stb'] } | Sort-Object -Unique)
}

$state = [ordered]@{
    CurrentObject = $project
    Journal = $journal
}

$json = $state | ConvertTo-Json -Depth 100
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[System.IO.File]::WriteAllText($outPath, $json, $utf8NoBom)

# Keep all project-local runtime copies in sync, so app startup never picks stale data.json.
$projectRoot = [System.IO.Path]::GetFullPath($PSScriptRoot)
$sourceBytes = [System.IO.File]::ReadAllBytes($outPath)
Get-ChildItem -Path $projectRoot -Recurse -File -Filter 'data.json' | ForEach-Object {
    $targetPath = $_.FullName
    if (-not [string]::Equals($targetPath, $outPath, [System.StringComparison]::CurrentCultureIgnoreCase)) {
        try {
            [System.IO.File]::WriteAllBytes($targetPath, $sourceBytes)
        }
        catch {
            # Ignore protected locations and keep going.
        }
    }
}

Write-Output "Generated test data: $outPath"
