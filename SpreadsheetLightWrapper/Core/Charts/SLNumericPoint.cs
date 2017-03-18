using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLightWrapper.Core.Charts
{
    internal class SLNumericPoint
    {
        internal SLNumericPoint()
        {
            NumericValue = string.Empty;
            Index = 0;
            FormatCode = string.Empty;
        }

        internal string NumericValue { get; set; }
        internal uint Index { get; set; }
        internal string FormatCode { get; set; }

        internal C.NumericPoint ToNumericPoint()
        {
            var np = new C.NumericPoint();
            np.Index = Index;
            if (FormatCode.Length > 0) np.FormatCode = FormatCode;
            np.NumericValue = new C.NumericValue(NumericValue);

            return np;
        }

        internal SLNumericPoint Clone()
        {
            var np = new SLNumericPoint();
            np.NumericValue = NumericValue;
            np.Index = Index;
            np.FormatCode = FormatCode;

            return np;
        }
    }
}