using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLightWrapper.Core.misc;

namespace SpreadsheetLightWrapper.Core.table
{
    internal class SLCalculatedColumnFormula
    {
        internal SLCalculatedColumnFormula()
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

        internal void FromCalculatedColumnFormula(CalculatedColumnFormula ccf)
        {
            SetAllNull();

            if ((ccf.Array != null) && ccf.Array.Value) Array = true;
            Text = ccf.Text;
        }

        internal CalculatedColumnFormula ToCalculatedColumnFormula()
        {
            var ccf = new CalculatedColumnFormula();
            if (Array) ccf.Array = Array;

            if (SLTool.ToPreserveSpace(Text))
                ccf.Space = SpaceProcessingModeValues.Preserve;
            ccf.Text = Text;

            return ccf;
        }

        internal SLCalculatedColumnFormula Clone()
        {
            var ccf = new SLCalculatedColumnFormula();
            ccf.Array = Array;
            ccf.Text = Text;

            return ccf;
        }
    }
}