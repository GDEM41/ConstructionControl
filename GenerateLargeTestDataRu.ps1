$ErrorActionPreference = 'Stop'

$path = Join-Path $PSScriptRoot 'data.json'
$jsonText = [System.IO.File]::ReadAllText($path, [System.Text.UTF8Encoding]::new($false))
$state = $jsonText | ConvertFrom-Json
if (-not $state.CurrentObject) { throw 'CurrentObject not found in data.json' }
$co = $state.CurrentObject

$startDate = [datetime]'2025-04-01'
$endDate = [datetime]'2026-03-31'

function RandInt([int]$min, [int]$max) { Get-Random -Minimum $min -Maximum ($max + 1) }
function Pick([object[]]$items) { return $items[(Get-Random -Minimum 0 -Maximum $items.Count)] }
function PickMany([object[]]$items, [int]$min, [int]$max) {
    $count = [Math]::Min($items.Count, (RandInt $min $max))
    return $items | Sort-Object { Get-Random } | Select-Object -First $count
}

$co.Name = 'Сокол Школка'
$co.BlocksCount = 3
$co.SameFloorsInBlocks = $true
$co.FloorsPerBlock = 4
$co.HasBasement = $true
$co.FloorsByBlock = @{}
$co.SummaryVisibleGroups = @()

$mainGroups = @('Ригели', 'Плиты перекрытия', 'Диафрагмы', 'Колонны')
$extraGroups = @('Внутренние работы', 'Малоценка', 'Инструмент', 'Расходники')
$allGroups = @($mainGroups + $extraGroups)
$marksMain = @{
    'Ригели' = @('+0.080', '+3.220', '+6.450')
    'Плиты перекрытия' = @('0.000', '+3.000', '+6.000')
    'Диафрагмы' = @('0.000', '+3.000', '+6.000')
    'Колонны' = @('0.000', '+3.300', '+6.600')
}

$materials = @{
    'Ригели' = @('РОП4.26-30', 'РО23-4', 'РДП4.56-50', 'РОП4.35-3', 'РГ-01', 'РГ-02', 'РГ-03', 'РГ-04', 'РГ-05', 'РГ-06')
    'Плиты перекрытия' = @('ПК56.15-12', 'ПК56.12-10', 'ПК56.15-10', 'ПК36.15-8', 'ПБ-1', 'ПБ-2', 'ПБ-3', 'ПБ-4', 'ПБ-5', 'ПБ-6')
    'Диафрагмы' = @('ДФ-1', 'ДФ-2', 'ДФ-3', 'ДФ-4', 'ДФ-5', 'ДФ-6', 'ДФ-7', 'ДФ-8', 'ДФ-9', 'ДФ-10')
    'Колонны' = @('К-1', 'К-2', 'К-3', 'К-4', 'К-5', 'К-6', 'К-7', 'К-8', 'К-9', 'К-10')
    'Внутренние работы' = @('Профиль ПС 50', 'ГКЛ 12.5', 'Шпаклевка', 'Грунтовка', 'Крепеж', 'Клей плиточный', 'Плитка', 'Кабель ВВГ', 'Автомат 16А', 'Розетка')
    'Малоценка' = @('Перчатки', 'Очки защитные', 'Каска', 'Маркер', 'Рулетка 5м', 'Лента сигнальная', 'Скотч', 'Кисть', 'Ведро', 'Мешки')
    'Инструмент' = @('Перфоратор', 'Болгарка', 'Шуруповерт', 'Лазерный уровень', 'Тачка', 'Лестница', 'Сварочный аппарат', 'Компрессор', 'Отбойный молоток', 'Пила')
    'Расходники' = @('Электроды', 'Диски отрезные', 'Бур 10мм', 'Бур 14мм', 'Сверло 6мм', 'Гвозди', 'Саморезы', 'Пена монтажная', 'Герметик', 'Изолента')
}

$co.MaterialNamesByGroup = @{}
$co.MaterialGroups = @()
foreach ($g in $allGroups) {
    $co.MaterialNamesByGroup[$g] = @($materials[$g])
    $co.MaterialGroups += [pscustomobject]@{
        Name = $g
        Items = @($materials[$g])
    }
}
$co.SummaryVisibleGroups = @($allGroups)

$co.SummaryMarksByGroup = @{}
foreach ($g in $mainGroups) { $co.SummaryMarksByGroup[$g] = @($marksMain[$g]) }
foreach ($g in $extraGroups) { $co.SummaryMarksByGroup[$g] = @() }

$co.StbByGroup = @{}
$co.SupplierByGroup = @{}
foreach ($g in $mainGroups) {
    $co.StbByGroup[$g] = 'СТБ 1300'
    $co.SupplierByGroup[$g] = 'ОАО МЖБ'
}
foreach ($g in $extraGroups) {
    $co.StbByGroup[$g] = 'ТУ'
    $co.SupplierByGroup[$g] = 'ООО Комплект'
}

$co.MaterialCatalog = @()
foreach ($g in $mainGroups) {
    foreach ($m in $materials[$g]) {
        $co.MaterialCatalog += [pscustomobject]@{
            CategoryName = 'Основные'
            TypeName = $g
            SubTypeName = ''
            ExtraLevels = @()
            LevelMarks = @($marksMain[$g])
            MaterialName = $m
        }
    }
}
foreach ($g in $extraGroups) {
    $type = if ($g -eq 'Малоценка') { 'Допы' } else { 'Внутренние' }
    foreach ($m in $materials[$g]) {
        $co.MaterialCatalog += [pscustomobject]@{
            CategoryName = 'Допы'
            TypeName = $type
            SubTypeName = $g
            ExtraLevels = @()
            LevelMarks = @()
            MaterialName = $m
        }
    }
}

$co.AutoSplitMaterialNames = @(
    'ПК56.15-12', 'ПК56.12-10', 'ПК56.15-10',
    'РОП4.26-30', 'РО23-4', 'РДП4.56-50'
)
$co.MaterialTreeSplitRules = @{
    'ПК56.15-12' = 'ПК|56|15|12'
    'ПК56.12-10' = 'ПК|56|12|10'
    'ПК56.15-10' = 'ПК|56|15|10'
    'РОП4.26-30' = 'РОП|4|26|30'
    'РО23-4' = 'РО|23|4'
    'РДП4.56-50' = 'РДП|4|56|50'
}

$co.Demand = @{}
$blocks = 1..$co.BlocksCount
foreach ($g in $mainGroups) {
    foreach ($m in $materials[$g]) {
        $levels = @{}
        $mounted = @{}
        foreach ($b in $blocks) {
            $lv = @{}
            $mv = @{}
            foreach ($mark in $marksMain[$g]) {
                $need = RandInt 12 48
                $done = RandInt 0 ($need - 1)
                $lv[$mark] = [double]$need
                $mv[$mark] = [double]$done
            }
            $levels["$b"] = $lv
            $mounted["$b"] = $mv
        }
        $co.Demand["$g::$m"] = [pscustomobject]@{
            Unit = 'шт'
            Levels = $levels
            MountedLevels = $mounted
            Floors = @{}
            MountedFloors = @{}
        }
    }
}

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

$co.UiSettings = [pscustomobject]@{
    DisableTree = $false
    PinTreeByDefault = $false
    ShowReminderPopup = $true
    ReminderSnoozeMinutes = 15
    HideReminderDetails = $false
}

$peopleSeed = @('Иванов Андрей Сергеевич','Петров Павел Алексеевич','Сидоров Илья Викторович','Смирнов Роман Николаевич','Волков Олег Дмитриевич','Егоров Кирилл Максимович','Козлов Антон Михайлович','Федоров Георгий Романович','Климов Артур Олегович','Руденко Сергей Павлович','Белый Дмитрий Игоревич','Мельник Алексей Андреевич','Горбунов Егор Сергеевич','Тихонов Илья Павлович','Жуков Никита Андреевич')
$specialties = @('Монтажник', 'Арматурщик', 'Бетонщик', 'Сварщик', 'Такелажник', 'Электрик')
$ranks = @('3', '4', '5', '6')
$brigades = 1..12 | ForEach-Object { "Бригада $_" }
$professions = @('Рабочий', 'Монтажник ЖБК', 'Арматурщик', 'Бетонщик', 'Электромонтажник')

$monthKeys = @()
$cursor = Get-Date -Date $startDate
while ($cursor -le $endDate) { $monthKeys += $cursor.ToString('yyyy-MM'); $cursor = $cursor.AddMonths(1) }

$co.TimesheetPeople = @()
$co.OtJournal = @()
for ($i = 0; $i -lt 120; $i++) {
    $name = "$($peopleSeed[$i % $peopleSeed.Count]) №$($i + 1)"
    $personId = [guid]::NewGuid().ToString()
    $brigade = $brigades[$i % $brigades.Count]
    $isBrigadier = (($i % 12) -eq 0)
    $months = @()
    foreach ($mk in $monthKeys) {
        $y = [int]$mk.Substring(0, 4); $m = [int]$mk.Substring(5, 2); $dim = [datetime]::DaysInMonth($y, $m)
        $dayValues = @{}; $dayEntries = @{}
        foreach ($day in 1..$dim) {
            $d = [datetime]::new($y, $m, $day)
            $isWeekend = $d.DayOfWeek -in @([System.DayOfWeek]::Saturday, [System.DayOfWeek]::Sunday)
            if ($isWeekend) { $value='В'; $comment=''; $doc=$null } else {
                $roll = RandInt 1 100
                if ($roll -le 82) { $value = Pick @('10','11','8'); $comment=''; $doc=$null } else { $value = Pick @('Б','О','К','Н'); $comment = Pick @('Больничный лист','Заявление на отпуск','Командировочное удостоверение','Справка'); $doc = ((RandInt 1 100) -le 70) }
            }
            $dayValues["$day"] = $value
            $dayEntries["$day"] = [pscustomobject]@{ Value = $value; PresenceMark = if ($value -match '^\d+$') { '✔' } else { '' }; Comment = $comment; DocumentAccepted = $doc; ArrivalMarked = ((RandInt 1 100) -le 14) }
        }
        $months += [pscustomobject]@{ MonthKey = $mk; DayValues = $dayValues; DayEntries = $dayEntries }
    }
    $co.TimesheetPeople += [pscustomobject]@{ PersonId = $personId; FullName = $name; Specialty = Pick $specialties; Rank = Pick $ranks; BrigadeName = $brigade; IsBrigadier = $isBrigadier; Months = $months }

    $primaryDate = $startDate.AddDays((RandInt 0 90))
    $insNo = RandInt 1 500
    $co.OtJournal += [pscustomobject]@{
        PersonId = $personId
        InstructionDate = $primaryDate.ToString('s')
        FullName = $name
        Specialty = Pick $specialties
        Rank = Pick $ranks
        Profession = Pick $professions
        InstructionType = 'Первичный на рабочем месте'
        InstructionNumbers = "ПР-$insNo"
        RepeatPeriodMonths = 3
        IsBrigadier = $isBrigadier
        BrigadierName = if ($isBrigadier) { '' } else { "$brigade (бригадир)" }
        IsDismissed = ((RandInt 1 100) -le 3)
        IsPendingRepeat = $false
        IsRepeatCompleted = $false
    }

    $repDate = $primaryDate.AddMonths(3)
    $repIdx = 1
    while ($repDate -le $endDate) {
        $pending = ((RandInt 1 100) -le 25)
        $co.OtJournal += [pscustomobject]@{
            PersonId = $personId
            InstructionDate = $repDate.ToString('s')
            FullName = $name
            Specialty = Pick $specialties
            Rank = Pick $ranks
            Profession = Pick $professions
            InstructionType = 'Повторный на рабочем месте'
            InstructionNumbers = "ПВ-$insNo-$repIdx"
            RepeatPeriodMonths = 3
            IsBrigadier = $isBrigadier
            BrigadierName = if ($isBrigadier) { '' } else { "$brigade (бригадир)" }
            IsDismissed = $false
            IsPendingRepeat = $pending
            IsRepeatCompleted = (-not $pending)
        }
        $repDate = $repDate.AddMonths(3)
        $repIdx++
    }
}

$actions = @('Монтаж', 'Кладка', 'Устройство')
$weatherTypes = @('ясно', 'облачно', 'пасмурно', 'дождь', 'снег')
$deviationTypes = @('Отклонений нет', 'Отклонение от разбивочных осей +3 мм', 'Отклонение от разбивочных осей +5 мм')
$co.ProductionJournal = @()
$dateCursor = $startDate
while ($dateCursor -le $endDate) {
    if ($dateCursor.DayOfWeek -notin @([System.DayOfWeek]::Saturday, [System.DayOfWeek]::Sunday)) {
        $rowsPerDay = RandInt 2 6
        for ($r = 0; $r -lt $rowsPerDay; $r++) {
            $g = Pick $mainGroups
            $elements = @(); foreach ($mat in (PickMany $materials[$g] 1 3)) { $elements += "$mat - $(RandInt 1 12)" }
            $blocksTxt = (PickMany @('1', '2', '3') 1 2) -join ', '
            $marksTxt = (PickMany $marksMain[$g] 1 2) -join ', '
            $temp = switch ($dateCursor.Month) { { $_ -in 12,1,2 } { RandInt -12 2 } { $_ -in 3,4 } { RandInt 0 12 } { $_ -in 5,6 } { RandInt 10 24 } { $_ -in 7,8 } { RandInt 18 32 } { $_ -in 9,10 } { RandInt 8 20 } default { RandInt -2 8 } }
            $remain = @(); foreach ($b in 1..3) { $mk = Pick $marksMain[$g]; $remain += "Блок $b, $mk: остаток $(RandInt 0 35)" }
            $co.ProductionJournal += [pscustomobject]@{ Date = $dateCursor.ToString('s'); ActionName = Pick $actions; WorkName = $g; ElementsText = ($elements -join '; '); BlocksText = $blocksTxt; MarksText = $marksTxt; BrigadeName = Pick $brigades; Weather = "$temp°C, $(Pick $weatherTypes)"; Deviations = Pick $deviationTypes; RequiresHiddenWorkAct = ((RandInt 1 100) -le 35); RemainingInfo = ($remain -join '; ') }
        }
    }
    $dateCursor = $dateCursor.AddDays(1)
}

$inspectionJournals = @('Журнал осмотра временных ограждений','Журнал осмотра лесов и подмостей','Журнал осмотра грузозахватных приспособлений','Журнал осмотра электроинструмента','Журнал осмотра средств пожаротушения','Журнал осмотра лестниц и стремянок')
$inspectionNames = @('Ежедневный осмотр','Еженедельный осмотр','Плановый осмотр перед сменой','Контроль состояния после работ')
$co.InspectionJournal = @()
for ($i = 0; $i -lt 120; $i++) {
    $start = $startDate.AddDays((RandInt 0 364)); $period = Pick @(7, 10, 14, 30); $lastDone = if ((RandInt 1 100) -le 82) { $start.AddDays((RandInt 0 300)) } else { $null }; if ($lastDone -and $lastDone -gt $endDate) { $lastDone = $endDate.AddDays(-(RandInt 0 20)) }
    $co.InspectionJournal += [pscustomobject]@{ JournalName = $inspectionJournals[$i % $inspectionJournals.Count]; InspectionName = "$($inspectionNames[$i % $inspectionNames.Count]) №$($i + 1)"; ReminderStartDate = $start.ToString('s'); ReminderPeriodDays = $period; LastCompletedDate = if ($lastDone) { $lastDone.ToString('s') } else { $null }; Notes = Pick @('Осмотр выполнен по чек-листу','Требуется подпись мастера','Без замечаний','Устранить замечания до конца смены') }
}

$state.Journal = @()
$ttnCounter = 1
$sheetName = 'Приход'
$dateCursor = $startDate
while ($dateCursor -le $endDate) {
    if ($dateCursor.DayOfWeek -notin @([System.DayOfWeek]::Saturday, [System.DayOfWeek]::Sunday)) {
        if ((RandInt 1 100) -le 95) {
            $arrivals = RandInt 2 6
            for ($a = 0; $a -lt $arrivals; $a++) {
                $isMain = ((RandInt 1 100) -le 72)
                if ($isMain) { $g = Pick $mainGroups; $cat = 'Основные'; $sub = ''; $unit = 'шт' } else { $g = Pick $extraGroups; $cat = 'Допы'; $sub = if ($g -eq 'Малоценка') { 'Малоценка' } else { 'Внутренние' }; $unit = if ($g -eq 'Малоценка' -or $g -eq 'Расходники') { 'кг' } else { 'шт' } }
                $rows = if ($isMain) { RandInt 3 7 } else { RandInt 2 4 }
                $selected = PickMany $materials[$g] $rows $rows
                $ttn = "{0}/{1}" -f $dateCursor.ToString('yyMMdd'), $ttnCounter
                $passport = "ПС-$($dateCursor.ToString('yyMMdd'))-$ttnCounter"
                foreach ($m in $selected) {
                    $qty = if ($unit -eq 'шт') { RandInt 2 24 } else { RandInt 20 650 }
                    $vol = if ($unit -eq 'шт') { [math]::Round($qty * 0.08, 2) } else { [math]::Round($qty / 1000.0, 2) }
                    $state.Journal += [pscustomobject]@{ SheetName = $sheetName; Date = $dateCursor.ToString('s'); ObjectName = $co.Name; Category = $cat; SubCategory = $sub; MaterialGroup = $g; MaterialName = $m; Unit = $unit; Quantity = [double]$qty; Passport = $passport; Ttn = $ttn; Stb = if ($isMain) { 'СТБ 1300' } else { 'ТУ' }; Supplier = if ($isMain) { 'ОАО МЖБ' } else { 'ООО Комплект' }; Position = "П-$($ttnCounter)-$(RandInt 1 99)"; Volume = $vol.ToString('0.##', [System.Globalization.CultureInfo]::InvariantCulture) }
                }
                $ttnCounter++
            }
        }
    }
    $dateCursor = $dateCursor.AddDays(1)
}

$archiveGroups = [ordered]@{}
foreach ($r in $state.Journal) { if (-not $archiveGroups.Contains($r.MaterialGroup)) { $archiveGroups[$r.MaterialGroup] = New-Object 'System.Collections.Generic.HashSet[string]' }; [void]$archiveGroups[$r.MaterialGroup].Add($r.MaterialName) }
$archiveMaterials = [ordered]@{}
foreach ($g in $archiveGroups.Keys) { $archiveMaterials[$g] = @($archiveGroups[$g] | Sort-Object) }
$co.Archive = [pscustomobject]@{ Groups = @($archiveGroups.Keys | Sort-Object); Materials = [pscustomobject]$archiveMaterials; Units = @($state.Journal | Select-Object -ExpandProperty Unit -Unique | Sort-Object); Suppliers = @($state.Journal | Select-Object -ExpandProperty Supplier -Unique | Sort-Object); Passports = @($state.Journal | Select-Object -ExpandProperty Passport -Unique | Sort-Object); Stb = @($state.Journal | Select-Object -ExpandProperty Stb -Unique | Sort-Object) }

$co.ArrivalHistory = @()
$co.PdfDocuments = @()
$co.EstimateDocuments = @()

$jsonOut = $state | ConvertTo-Json -Depth 100
[System.IO.File]::WriteAllText($path, $jsonOut, [System.Text.UTF8Encoding]::new($false))
Write-Output 'Русская тестовая база сгенерирована.'
