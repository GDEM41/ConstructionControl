using System;
using System.Collections.Generic;

namespace ConstructionControl
{
    [Flags]
    public enum ProjectTransferSection
    {
        None = 0,
        ObjectSettings = 1 << 0,
        MaterialsAndSummary = 1 << 1,
        Arrival = 1 << 2,
        Ot = 1 << 3,
        Timesheet = 1 << 4,
        Production = 1 << 5,
        HiddenWorkActs = 1 << 6,
        Inspection = 1 << 7,
        Notes = 1 << 8,
        Pdf = 1 << 9,
        Estimates = 1 << 10
    }

    public sealed class ProjectSectionSelectionOption
    {
        public ProjectTransferSection Section { get; set; }
        public string Title { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public bool IsSelected { get; set; }
        public bool IsEnabled { get; set; } = true;
    }

    public sealed class BackupPackageManifest
    {
        public int Version { get; set; } = 1;
        public List<string> IncludedSections { get; set; } = new();
    }
}
