using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.table
{
    internal class SLTop10
    {
        internal SLTop10()
        {
            SetAllNull();
        }

        internal bool? Top { get; set; }
        internal bool? Percent { get; set; }
        internal double Val { get; set; }
        internal double? FilterValue { get; set; }

        private void SetAllNull()
        {
            Top = null;
            Percent = null;
            Val = 0.0;
            FilterValue = null;
        }

        internal void FromTop10(Top10 t)
        {
            SetAllNull();

            if (t.Top != null) Top = t.Top.Value;
            if (t.Percent != null) Percent = t.Percent.Value;
            Val = t.Val.Value;
            if (t.FilterValue != null) FilterValue = t.FilterValue.Value;
        }

        internal Top10 ToTop10()
        {
            var t = new Top10();
            if ((Top != null) && !Top.Value) t.Top = Top.Value;
            if ((Percent != null) && Percent.Value) t.Percent = Percent.Value;
            t.Val = Val;
            if (FilterValue != null) t.FilterValue = FilterValue.Value;

            return t;
        }

        internal SLTop10 Clone()
        {
            var t = new SLTop10();
            t.Top = Top;
            t.Percent = Percent;
            t.Val = Val;
            t.FilterValue = FilterValue;

            return t;
        }
    }
}