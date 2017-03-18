using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.pivottable
{
    internal class SLCalculatedItem
    {
        internal SLCalculatedItem()
        {
            SetAllNull();
        }

        internal SLPivotArea PivotArea { get; set; }

        internal uint? Field { get; set; }
        internal string Formula { get; set; }

        private void SetAllNull()
        {
            PivotArea = new SLPivotArea();
            Field = null;
            Formula = "";
        }

        internal void FromCalculatedItem(CalculatedItem ci)
        {
            SetAllNull();

            if (ci.Field != null) Field = ci.Field.Value;
            if (ci.Formula != null) Formula = ci.Formula.Value;

            if (ci.PivotArea != null) PivotArea.FromPivotArea(ci.PivotArea);
        }

        internal CalculatedItem ToCalculatedItem()
        {
            var ci = new CalculatedItem();
            if (Field != null) ci.Field = Field.Value;
            if ((Formula != null) && (Formula.Length > 0)) ci.Formula = Formula;

            ci.PivotArea = PivotArea.ToPivotArea();

            return ci;
        }

        internal SLCalculatedItem Clone()
        {
            var ci = new SLCalculatedItem();
            ci.Field = Field;
            ci.Formula = Formula;
            ci.PivotArea = PivotArea.Clone();

            return ci;
        }
    }
}