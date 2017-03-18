using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.pivottable
{
    internal class SLKpi
    {
        // what happened to Time attribute?

        internal SLKpi()
        {
            SetAllNull();
        }

        internal string UniqueName { get; set; }
        internal string Caption { get; set; }
        internal string DisplayFolder { get; set; }
        internal string MeasureGroup { get; set; }
        internal string ParentKpi { get; set; }
        internal string Value { get; set; }
        internal string Goal { get; set; }
        internal string Status { get; set; }
        internal string Trend { get; set; }
        internal string Weight { get; set; }

        private void SetAllNull()
        {
            UniqueName = "";
            Caption = "";
            DisplayFolder = "";
            MeasureGroup = "";
            ParentKpi = "";
            Value = "";
            Goal = "";
            Status = "";
            Trend = "";
            Weight = "";
        }

        internal void FromKpi(Kpi k)
        {
            SetAllNull();

            if (k.UniqueName != null) UniqueName = k.UniqueName.Value;
            if (k.Caption != null) Caption = k.Caption.Value;
            if (k.DisplayFolder != null) DisplayFolder = k.DisplayFolder.Value;
            if (k.MeasureGroup != null) MeasureGroup = k.MeasureGroup.Value;
            if (k.ParentKpi != null) ParentKpi = k.ParentKpi.Value;
            if (k.Value != null) Value = k.Value.Value;
            if (k.Goal != null) Goal = k.Goal.Value;
            if (k.Status != null) Status = k.Status.Value;
            if (k.Trend != null) Trend = k.Trend.Value;
            if (k.Weight != null) Weight = k.Weight.Value;
        }

        internal Kpi ToKpi()
        {
            var k = new Kpi();
            k.UniqueName = UniqueName;
            if ((Caption != null) && (Caption.Length > 0)) k.Caption = Caption;
            if ((DisplayFolder != null) && (DisplayFolder.Length > 0)) k.DisplayFolder = DisplayFolder;
            if ((MeasureGroup != null) && (MeasureGroup.Length > 0)) k.MeasureGroup = MeasureGroup;
            if ((ParentKpi != null) && (ParentKpi.Length > 0)) k.ParentKpi = ParentKpi;
            k.Value = Value;
            if ((Goal != null) && (Goal.Length > 0)) k.Goal = Goal;
            if ((Status != null) && (Status.Length > 0)) k.Status = Status;
            if ((Trend != null) && (Trend.Length > 0)) k.Trend = Trend;
            if ((Weight != null) && (Weight.Length > 0)) k.Weight = Weight;

            return k;
        }

        internal SLKpi Clone()
        {
            var k = new SLKpi();
            k.UniqueName = UniqueName;
            k.Caption = Caption;
            k.DisplayFolder = DisplayFolder;
            k.MeasureGroup = MeasureGroup;
            k.ParentKpi = ParentKpi;
            k.Value = Value;
            k.Goal = Goal;
            k.Status = Status;
            k.Trend = Trend;
            k.Weight = Weight;

            return k;
        }
    }
}