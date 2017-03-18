using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Ups.Toolkit.SpreadsheetLight.Core.misc;

namespace Ups.Toolkit.SpreadsheetLight.Core.table
{
    internal class SLTotalsRowFormula
    {
        internal SLTotalsRowFormula()
        {
            SetAllNull();
        }

        internal bool Array { get; set; }
        internal string Text { get; set; }

        private void SetAllNull()
        {
            Array = false;
            Text = string.Empty;
        }

        internal void FromTotalsRowFormula(TotalsRowFormula trf)
        {
            SetAllNull();

            if ((trf.Array != null) && trf.Array.Value) Array = true;
            Text = trf.Text;
        }

        internal TotalsRowFormula ToTotalsRowFormula()
        {
            var trf = new TotalsRowFormula();
            if (Array) trf.Array = Array;

            if (SLTool.ToPreserveSpace(Text))
                trf.Space = SpaceProcessingModeValues.Preserve;
            trf.Text = Text;

            return trf;
        }

        internal SLTotalsRowFormula Clone()
        {
            var trf = new SLTotalsRowFormula();
            trf.Array = Array;
            trf.Text = Text;

            return trf;
        }
    }
}