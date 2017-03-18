using System.Collections.Generic;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLightWrapper.Core.Charts
{
    internal class SLMultiLevelStringReference
    {
        internal SLMultiLevelStringReference()
        {
            Formula = string.Empty;
            MultiLevelStringCache = new SLMultiLevelStringCache();
        }

        internal string Formula { get; set; }
        internal SLMultiLevelStringCache MultiLevelStringCache { get; set; }

        internal C.MultiLevelStringReference ToMultiLevelStringReference()
        {
            var mlsr = new C.MultiLevelStringReference();
            mlsr.Formula = new C.Formula(Formula);
            mlsr.MultiLevelStringCache = MultiLevelStringCache.ToMultiLevelStringCache();

            return mlsr;
        }

        internal SLMultiLevelStringReference Clone()
        {
            var mlsr = new SLMultiLevelStringReference();
            mlsr.Formula = Formula;
            mlsr.MultiLevelStringCache = MultiLevelStringCache.Clone();

            return mlsr;
        }
    }

    internal class SLMultiLevelStringCache
    {
        internal SLMultiLevelStringCache()
        {
            PointCount = 0;
            Levels = new List<SLLevel>();
        }

        internal uint PointCount { get; set; }

        internal List<SLLevel> Levels { get; set; }

        internal C.MultiLevelStringCache ToMultiLevelStringCache()
        {
            var mlsc = new C.MultiLevelStringCache();
            mlsc.PointCount = new C.PointCount {Val = PointCount};

            C.Level lvl;
            int i, j;
            for (i = 0; i < Levels.Count; ++i)
            {
                lvl = new C.Level();
                for (j = 0; j < Levels[i].Points.Count; ++j)
                    lvl.Append(Levels[i].Points[j].ToStringPoint());
                mlsc.Append(lvl);
            }

            return mlsc;
        }

        internal SLMultiLevelStringCache Clone()
        {
            var mlsc = new SLMultiLevelStringCache();
            mlsc.PointCount = PointCount;
            for (var i = 0; i < Levels.Count; ++i)
                mlsc.Levels.Add(Levels[i].Clone());

            return mlsc;
        }
    }

    internal class SLLevel
    {
        internal SLLevel()
        {
            Points = new List<SLStringPoint>();
        }

        internal List<SLStringPoint> Points { get; set; }

        internal SLLevel Clone()
        {
            var lvl = new SLLevel();
            for (var i = 0; i < Points.Count; ++i)
                lvl.Points.Add(Points[i].Clone());

            return lvl;
        }
    }
}