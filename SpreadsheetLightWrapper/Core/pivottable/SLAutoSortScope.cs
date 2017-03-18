using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.pivottable
{
    internal class SLAutoSortScope
    {
        internal SLAutoSortScope()
        {
            SetAllNull();
        }

        internal SLPivotArea PivotArea { get; set; }

        private void SetAllNull()
        {
            PivotArea = new SLPivotArea();
        }

        // ahahahahah... I did *not* just come up with this variable name... :)
        internal void FromAutoSortScope(AutoSortScope ass)
        {
            SetAllNull();

            if (ass.PivotArea != null) PivotArea.FromPivotArea(ass.PivotArea);
        }

        internal AutoSortScope ToAutoSortScope()
        {
            var ass = new AutoSortScope();
            ass.PivotArea = PivotArea.ToPivotArea();

            return ass;
        }

        internal SLAutoSortScope Clone()
        {
            var ass = new SLAutoSortScope();
            ass.PivotArea = PivotArea.Clone();

            return ass;
        }
    }
}