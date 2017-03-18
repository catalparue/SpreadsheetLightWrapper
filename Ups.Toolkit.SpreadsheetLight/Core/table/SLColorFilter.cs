using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.table
{
    internal class SLColorFilter
    {
        internal SLColorFilter()
        {
            SetAllNull();
        }

        internal uint? FormatId { get; set; }
        internal bool? CellColor { get; set; }

        private void SetAllNull()
        {
            FormatId = null;
            CellColor = null;
        }

        internal void FromColorFilter(ColorFilter cf)
        {
            SetAllNull();

            if (cf.FormatId != null) FormatId = cf.FormatId.Value;
            if ((cf.CellColor != null) && !cf.CellColor.Value) CellColor = cf.CellColor.Value;
        }

        internal ColorFilter ToColorFilter()
        {
            var cf = new ColorFilter();
            if (FormatId != null) cf.FormatId = FormatId.Value;
            if ((CellColor != null) && !CellColor.Value) cf.CellColor = CellColor.Value;

            return cf;
        }

        internal SLColorFilter Clone()
        {
            var cf = new SLColorFilter();
            cf.FormatId = FormatId;
            cf.CellColor = CellColor;

            return cf;
        }
    }
}