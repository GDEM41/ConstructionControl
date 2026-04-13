using System;
using System.Collections.Generic;

namespace ConstructionControl
{
    internal static class AppStateMigration
    {
        public static void Apply(AppState state)
        {
            if (state == null)
                return;

            if (state.SchemaVersion <= 0)
                state.SchemaVersion = 1;

            if (state.SchemaVersion < 2)
            {
                state.CurrentObject ??= new ProjectObject();
                state.CurrentObject.ChangeLog ??= new List<ProjectChangeLogEntry>();
                state.SchemaVersion = 2;
            }

            state.SavedAtUtc = state.SavedAtUtc == default ? DateTime.UtcNow : state.SavedAtUtc;
            state.Journal ??= new List<JournalRecord>();
            state.CurrentObject ??= new ProjectObject();
            state.CurrentObject.ChangeLog ??= new List<ProjectChangeLogEntry>();
            state.CurrentObject.UiSettings ??= new ProjectUiSettings();
            if (string.IsNullOrWhiteSpace(state.CurrentObject.UiSettings.AccessRole))
            {
                state.CurrentObject.UiSettings.AccessRole = ProjectAccessRoles.Critical;
                if (!state.CurrentObject.UiSettings.RequireCodeForCriticalOperations)
                    state.CurrentObject.UiSettings.RequireCodeForCriticalOperations = true;
            }
        }
    }
}
