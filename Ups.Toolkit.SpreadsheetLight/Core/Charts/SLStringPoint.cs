using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Ups.Toolkit.SpreadsheetLight.Core.Charts
{
    internal class SLStringPoint
    {
        internal SLStringPoint()
        {
            NumericValue = string.Empty;
            Index = 0;
        }

        internal string NumericValue { get; set; }
        internal uint Index { get; set; }

        internal C.StringPoint ToStringPoint()
        {
            var sp = new C.StringPoint();
            sp.Index = Index;
            sp.NumericValue = new C.NumericValue(NumericValue);

            return sp;
        }

        internal SLStringPoint Clone()
        {
            var sp = new SLStringPoint();
            sp.NumericValue = NumericValue;
            sp.Index = Index;

            return sp;
        }
    }
}