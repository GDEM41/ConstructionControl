using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace ConstructionControl
{
    public class ProjectObject
    {
        public string Name { get; set; }

        // ===== НОВЫЕ НАСТРОЙКИ ОБЪЕКТА =====

        // Количество блоков
        public int BlocksCount { get; set; }

        // Есть ли подвал
        public bool HasBasement { get; set; }

        // Одинаковое количество этажей во всех блоках
        public bool SameFloorsInBlocks { get; set; } = true;

        // Если этажи одинаковые
        public int FloorsPerBlock { get; set; }

        // Если этажи разные (ключ = номер блока, значение = этажи)
        public Dictionary<int, int> FloorsByBlock { get; set; } = new();

        // ===== СТАРОЕ (НЕ ТРОГАЕМ) =====

        public Dictionary<string, List<string>> MaterialNamesByGroup { get; set; } = new();

        public Dictionary<string, string> StbByGroup { get; set; } = new();
        public Dictionary<string, string> SupplierByGroup { get; set; } = new();

        public List<MaterialGroup> MaterialGroups { get; set; } = new();
        public List<ArrivalItem> ArrivalHistory { get; set; } = new();
    }


    public class MaterialGroup
    {
        public string Name { get; set; }
        public List<string> Items { get; set; } = new();
    }

    public class Arrival
    {
        public string MaterialGroup { get; set; }
        public string TtnNumber { get; set; }
        public List<ArrivalItem> Items { get; set; } = new();
    }

    public class ArrivalItem
    {
        public DateTime Date { get; set; } = DateTime.Today;
        public string MaterialName { get; set; }
        public string Unit { get; set; }
        public int Quantity { get; set; }
        public string Passport { get; set; }
        public string Stb { get; set; }
        public string Supplier { get; set; }

        public ObservableCollection<string> AvailableNames { get; set; } = new();
        public ObservableCollection<string> AvailableUnits { get; set; } = new();
    }

    public class JournalRecord
    {
        public DateTime Date { get; set; }
        public string ObjectName { get; set; }
        public string MaterialGroup { get; set; }
        public string MaterialName { get; set; }
        public string Unit { get; set; }
        public int Quantity { get; set; }
        public string Passport { get; set; }
        public string Ttn { get; set; }
        public string Stb { get; set; }
        public string Supplier { get; set; }
    }
}
