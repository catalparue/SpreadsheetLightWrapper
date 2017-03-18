using System.Collections.Generic;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLightWrapper.Core.Charts
{
    internal class SLNumberingCache : SLNumberDataType
    {
        internal SLNumberingCache Clone()
        {
            var nc = new SLNumberingCache();
            nc.FormatCode = FormatCode;
            nc.PointCount = PointCount;
            for (var i = 0; i < Points.Count; ++i)
                nc.Points.Add(Points[i].Clone());

            return nc;
        }
    }

    internal class SLNumberLiteral : SLNumberDataType
    {
        internal SLNumberLiteral Clone()
        {
            var nl = new SLNumberLiteral();
            nl.FormatCode = FormatCode;
            nl.PointCount = PointCount;
            for (var i = 0; i < Points.Count; ++i)
                nl.Points.Add(Points[i].Clone());

            return nl;
        }
    }

    /// <summary>
    ///     For NumberingCache and NumberLiteral
    /// </summary>
    internal abstract class SLNumberDataType
    {
        internal SLNumberDataType()
        {
            FormatCode = string.Empty;
            PointCount = 0;
            Points = new List<SLNumericPoint>();
        }

        internal string FormatCode { get; set; }
        internal uint PointCount { get; set; }
        internal List<SLNumericPoint> Points { get; set; }

        internal C.NumberingCache ToNumberingCache()
        {
            var nc = new C.NumberingCache();
            nc.FormatCode = new C.FormatCode(FormatCode);
            nc.PointCount = new C.PointCount {Val = PointCount};
            for (var i = 0; i < Points.Count; ++i)
                nc.Append(Points[i].ToNumericPoint());

            return nc;
        }

        internal C.NumberLiteral ToNumberLiteral()
        {
            var nl = new C.NumberLiteral();
            nl.FormatCode = new C.FormatCode(FormatCode);
            nl.PointCount = new C.PointCount {Val = PointCount};
            for (var i = 0; i < Points.Count; ++i)
                nl.Append(Points[i].ToNumericPoint());

            return nl;
        }
    }
}