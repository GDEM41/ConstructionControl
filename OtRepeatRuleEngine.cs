using System;
using System.Collections.Generic;
using System.Linq;

namespace ConstructionControl
{
    internal static class OtRepeatRuleEngine
    {
        private static bool IsRepeatInstruction(OtJournalEntry entry)
            => !string.IsNullOrWhiteSpace(entry?.InstructionType)
                && entry.InstructionType.Contains("повторн", StringComparison.CurrentCultureIgnoreCase);

        public static List<OtJournalEntry> Apply(
            IList<OtJournalEntry> rows,
            DateTime today,
            Func<int, string> buildRepeatInstructionType)
        {
            var toAdd = new List<OtJournalEntry>();
            if (rows == null || rows.Count == 0)
                return toAdd;

            foreach (var scheduledRow in rows
                         .Where(x => x != null && !x.IsDismissed && x.IsScheduledRepeat && today.Date >= x.InstructionDate.Date)
                         .ToList())
            {
                scheduledRow.IsScheduledRepeat = false;
                scheduledRow.IsPendingRepeat = true;
                scheduledRow.IsRepeatCompleted = false;
            }

            foreach (var group in rows
                         .Where(x => x != null && !string.IsNullOrWhiteSpace(x.FullName))
                         .GroupBy(x => x.FullName.Trim(), StringComparer.CurrentCultureIgnoreCase))
            {
                var activeRows = group.Where(x => !x.IsDismissed).ToList();
                if (activeRows.Count == 0)
                    continue;

                if (activeRows.Any(x => x.IsPendingRepeat))
                    continue;

                var lastCompleted = activeRows
                    .Where(x => !x.IsPendingRepeat && !x.IsScheduledRepeat)
                    .OrderByDescending(x => x.InstructionDate)
                    .FirstOrDefault();

                if (lastCompleted == null)
                    continue;

                var repeatDate = lastCompleted.NextRepeatDate.Date;
                if (today.Date < repeatDate)
                    continue;

                var repeatIndex = group.Count(IsRepeatInstruction) + 1;
                var instructionType = buildRepeatInstructionType != null
                    ? buildRepeatInstructionType(repeatIndex)
                    : (repeatIndex <= 1 ? "Повторный" : $"Повторный ({repeatIndex})");

                var clone = new OtJournalEntry
                {
                    PersonId = lastCompleted.PersonId,
                    InstructionDate = repeatDate,
                    FullName = lastCompleted.FullName,
                    Specialty = lastCompleted.Specialty,
                    Rank = lastCompleted.Rank,
                    Profession = lastCompleted.Profession,
                    InstructionType = instructionType,
                    InstructionNumbers = lastCompleted.InstructionNumbers,
                    RepeatPeriodMonths = Math.Max(1, lastCompleted.RepeatPeriodMonths),
                    IsBrigadier = lastCompleted.IsBrigadier,
                    BrigadierName = lastCompleted.BrigadierName,
                    IsPendingRepeat = true,
                    IsScheduledRepeat = false,
                    IsRepeatCompleted = false,
                    IsDismissed = false
                };

                toAdd.Add(clone);
            }

            return toAdd;
        }
    }
}
