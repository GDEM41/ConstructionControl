using ConstructionControl;

var failures = new List<string>();

Run("Summary overage rule", () =>
{
    Assert(CalculationCore.ShouldNotifySummaryDelta(15, 10, includeOverage: true, includeDeficit: false), "Ожидалось уведомление о переходе.");
    Assert(!CalculationCore.ShouldNotifySummaryDelta(8, 10, includeOverage: true, includeDeficit: false), "Недоход не должен попадать в overage-режим.");
});

Run("Summary deficit rule", () =>
{
    Assert(CalculationCore.ShouldNotifySummaryDelta(8, 10, includeOverage: false, includeDeficit: true), "Ожидалось уведомление о недоходе.");
    Assert(!CalculationCore.ShouldNotifySummaryDelta(12, 10, includeOverage: false, includeDeficit: true), "Переход не должен попадать в deficit-режим.");
});

Run("Production clamp rule", () =>
{
    var clamped = CalculationCore.ClampToAvailable(requested: 9, arrived: 14, mounted: 8);
    Assert(Math.Abs(clamped - 6) < 0.0001, $"Ожидалось 6, получено {clamped:0.###}");
});

Run("OT repeat engine", () =>
{
    var today = new DateTime(2026, 4, 12);
    var rows = new List<OtJournalEntry>
    {
        new()
        {
            PersonId = Guid.NewGuid(),
            FullName = "Иванов Иван Иванович",
            InstructionDate = new DateTime(2025, 12, 1),
            InstructionType = "Первичный на рабочем месте",
            RepeatPeriodMonths = 3,
            IsDismissed = false,
            IsPendingRepeat = false,
            IsScheduledRepeat = false,
            IsRepeatCompleted = false
        }
    };

    var added = OtRepeatRuleEngine.Apply(rows, today, index => index == 1 ? "Повторный" : $"Повторный ({index})");
    Assert(added.Count == 1, $"Ожидалась одна запись, получено {added.Count}.");
    Assert(added[0].IsPendingRepeat, "Новая запись должна требовать повторный инструктаж.");
});

Run("State migration to latest schema", () =>
{
    var state = new AppState
    {
        SchemaVersion = 0,
        SavedAtUtc = default,
        CurrentObject = null,
        Journal = null
    };

    AppStateMigration.Apply(state);
    Assert(state.SchemaVersion == AppState.LatestSchemaVersion, "Версия схемы не обновилась.");
    Assert(state.CurrentObject != null, "CurrentObject должен быть инициализирован.");
    Assert(state.Journal != null, "Журнал должен быть инициализирован.");
    Assert(state.CurrentObject.UiSettings != null, "UiSettings должны быть инициализированы.");
});

if (failures.Count > 0)
{
    Console.Error.WriteLine("Автотесты завершились с ошибками:");
    foreach (var failure in failures)
        Console.Error.WriteLine($" - {failure}");
    Environment.Exit(1);
}

Console.WriteLine("Автотесты ключевой логики пройдены.");

return;

void Run(string name, Action test)
{
    try
    {
        test();
        Console.WriteLine($"[OK] {name}");
    }
    catch (Exception ex)
    {
        failures.Add($"{name}: {ex.Message}");
        Console.Error.WriteLine($"[FAIL] {name}: {ex.Message}");
    }
}

static void Assert(bool condition, string message)
{
    if (!condition)
        throw new InvalidOperationException(message);
}
