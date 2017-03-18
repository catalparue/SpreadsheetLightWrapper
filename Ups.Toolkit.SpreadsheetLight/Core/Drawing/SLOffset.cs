using A = DocumentFormat.OpenXml.Drawing;

namespace Ups.Toolkit.SpreadsheetLight.Core.Drawing
{
    internal class SLOffset
    {
        internal SLOffset()
        {
            X = 0;
            Y = 0;
        }

        internal long X { get; set; }
        internal long Y { get; set; }

        internal A.Offset ToOffset()
        {
            var off = new A.Offset();
            off.X = X;
            off.Y = Y;

            return off;
        }

        internal SLOffset Clone()
        {
            var off = new SLOffset();
            off.X = X;
            off.Y = Y;

            return off;
        }
    }
}