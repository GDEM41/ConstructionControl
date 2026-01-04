namespace ConstructionControl
{
    public class ExcelImportTemplate
    {
        public int DateRow { get; set; }
        public int MaterialColumn { get; set; }
        public int QuantityStartColumn { get; set; }

        public int? PositionColumn { get; set; }
        public int? UnitColumn { get; set; }
        public int? VolumeColumn { get; set; }
        public int? StbColumn { get; set; }

        public int? TtnRow { get; set; }
        public int? SupplierRow { get; set; }
        public int? PassportRow { get; set; }
    }
}
