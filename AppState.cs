using System.Collections.Generic;

using System;

namespace ConstructionControl
{
    public class AppState
    {
        public const int LatestSchemaVersion = 2;
        public int SchemaVersion { get; set; } = LatestSchemaVersion;
        public DateTime SavedAtUtc { get; set; } = DateTime.UtcNow;
        public ProjectObject? CurrentObject { get; set; }
        public List<JournalRecord> Journal { get; set; } = new();
    }
}
