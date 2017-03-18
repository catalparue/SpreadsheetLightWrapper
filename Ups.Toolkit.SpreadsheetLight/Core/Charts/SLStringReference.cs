using System.Collections.Generic;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Ups.Toolkit.SpreadsheetLight.Core.Charts
{
    internal class SLStringReference
    {
        internal SLStringReference()
        {
            WorksheetName = string.Empty;
            StartRowIndex = 1;
            StartColumnIndex = 1;
            EndRowIndex = 1;
            EndColumnIndex = 1;

            Formula = string.Empty;
            PointCount = 0;
            Points = new List<SLStringPoint>();
        }

        internal string WorksheetName { get; set; }
        internal int StartRowIndex { get; set; }
        internal int StartColumnIndex { get; set; }
        internal int EndRowIndex { get; set; }
        internal int EndColumnIndex { get; set; }

        internal string Formula { get; set; }

        // this is StringCache
        internal uint PointCount { get; set; }

        /// <summary>
        ///     This takes the place of StringCache
        /// </summary>
        internal List<SLStringPoint> Points { get; set; }

        internal C.StringReference ToStringReference()
        {
            var sr = new C.StringReference();
            sr.Formula = new C.Formula(Formula);
            sr.StringCache = new C.StringCache();
            sr.StringCache.PointCount = new C.PointCount {Val = PointCount};
            for (var i = 0; i < Points.Count; ++i)
                sr.StringCache.Append(Points[i].ToStringPoint());

            return sr;
        }

        internal void RefreshFormula()
        {
            Formula = SLChartTool.GetChartReferenceFormula(WorksheetName, StartRowIndex, StartColumnIndex, EndRowIndex,
                EndColumnIndex);
        }

        internal SLStringReference Clone()
        {
            var sr = new SLStringReference();
            sr.WorksheetName = WorksheetName;
            sr.StartRowIndex = StartRowIndex;
            sr.StartColumnIndex = StartColumnIndex;
            sr.EndRowIndex = EndRowIndex;
            sr.EndColumnIndex = EndColumnIndex;
            sr.Formula = Formula;
            sr.PointCount = PointCount;
            for (var i = 0; i < Points.Count; ++i)
                sr.Points.Add(Points[i].Clone());

            return sr;
        }
    }
}