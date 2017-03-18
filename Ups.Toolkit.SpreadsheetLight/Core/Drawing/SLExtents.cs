using A = DocumentFormat.OpenXml.Drawing;

namespace Ups.Toolkit.SpreadsheetLight.Core.Drawing
{
    internal class SLExtents
    {
        internal SLExtents()
        {
            Cx = 0;
            Cy = 0;
        }

        internal long Cx { get; set; }
        internal long Cy { get; set; }

        internal A.Extents ToExtents()
        {
            var ext = new A.Extents();
            ext.Cx = Cx;
            ext.Cy = Cy;

            return ext;
        }

        internal SLExtents Clone()
        {
            var ext = new SLExtents();
            ext.Cx = Cx;
            ext.Cy = Cy;

            return ext;
        }
    }
}