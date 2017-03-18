using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.pivottable
{
    internal class SLFormat
    {
        internal SLFormat()
        {
            SetAllNull();
        }

        internal SLPivotArea PivotArea { get; set; }
        internal FormatActionValues Action { get; set; }
        internal uint? FormatId { get; set; }

        private void SetAllNull()
        {
            PivotArea = new SLPivotArea();
            Action = FormatActionValues.Formatting;
            FormatId = null;
        }

        internal void FromFormat(Format f)
        {
            SetAllNull();

            if (f.PivotArea != null) PivotArea.FromPivotArea(f.PivotArea);

            if (f.Action != null) Action = f.Action.Value;
            if (f.FormatId != null) FormatId = f.FormatId.Value;
        }

        internal Format ToFormat()
        {
            var f = new Format();
            f.PivotArea = PivotArea.ToPivotArea();

            if (Action != FormatActionValues.Formatting) f.Action = Action;
            if (FormatId != null) f.FormatId = FormatId.Value;

            return f;
        }

        internal SLFormat Clone()
        {
            var f = new SLFormat();
            f.PivotArea = PivotArea.Clone();

            f.Action = Action;
            f.FormatId = FormatId;

            return f;
        }
    }
}