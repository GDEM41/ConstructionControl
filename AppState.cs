using System.Collections.Generic;

namespace ConstructionControl
{
    public class AppState
    {
        public ProjectObject CurrentObject { get; set; }
        public List<JournalRecord> Journal { get; set; } = new();
    }
}
