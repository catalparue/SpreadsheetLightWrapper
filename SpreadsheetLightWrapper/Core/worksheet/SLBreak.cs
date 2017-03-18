using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.worksheet
{
    internal class SLBreak
    {
        internal SLBreak()
        {
            SetAllNull();
        }

        internal uint Id { get; set; }
        internal uint Min { get; set; }
        internal uint Max { get; set; }
        internal bool ManualPageBreak { get; set; }
        internal bool PivotTablePageBreak { get; set; }

        internal void SetAllNull()
        {
            Id = 0;
            Min = 0;
            Max = 0;
            ManualPageBreak = false;
            PivotTablePageBreak = false;
        }

        internal void FromBreak(Break b)
        {
            SetAllNull();
            if (b.Id != null) Id = b.Id;
            if (b.Min != null) Min = b.Min;
            if (b.Max != null) Max = b.Max;
            if (b.ManualPageBreak != null) ManualPageBreak = b.ManualPageBreak;
            if (b.PivotTablePageBreak != null) PivotTablePageBreak = b.PivotTablePageBreak;
        }

        internal Break ToBreak()
        {
            var b = new Break();
            if (Id != 0) b.Id = Id;
            if (Min != 0) b.Min = Min;
            if (Max != 0) b.Max = Max;
            if (ManualPageBreak) b.ManualPageBreak = ManualPageBreak;
            if (PivotTablePageBreak) b.PivotTablePageBreak = PivotTablePageBreak;

            return b;
        }

        internal SLBreak Clone()
        {
            var b = new SLBreak();
            b.Id = Id;
            b.Min = Min;
            b.Max = Max;
            b.ManualPageBreak = ManualPageBreak;
            b.PivotTablePageBreak = PivotTablePageBreak;

            return b;
        }
    }
}