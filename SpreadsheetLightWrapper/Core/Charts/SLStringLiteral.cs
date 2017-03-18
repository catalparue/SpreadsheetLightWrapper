using System.Collections.Generic;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLightWrapper.Core.Charts
{
    internal class SLStringLiteral
    {
        internal SLStringLiteral()
        {
            PointCount = 0;
            Points = new List<SLStringPoint>();
        }

        internal uint PointCount { get; set; }
        internal List<SLStringPoint> Points { get; set; }

        internal C.StringLiteral ToStringLiteral()
        {
            var sl = new C.StringLiteral();
            sl.PointCount = new C.PointCount {Val = PointCount};
            for (var i = 0; i < Points.Count; ++i)
                sl.Append(Points[i].ToStringPoint());

            return sl;
        }

        internal SLStringLiteral Clone()
        {
            var sl = new SLStringLiteral();
            sl.PointCount = PointCount;
            for (var i = 0; i < Points.Count; ++i)
                sl.Points.Add(Points[i].Clone());

            return sl;
        }
    }
}