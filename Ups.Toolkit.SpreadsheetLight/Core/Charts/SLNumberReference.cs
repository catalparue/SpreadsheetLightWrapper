using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Ups.Toolkit.SpreadsheetLight.Core.Charts
{
    internal class SLNumberReference
    {
        internal SLNumberReference()
        {
            WorksheetName = string.Empty;
            StartRowIndex = 1;
            StartColumnIndex = 1;
            EndRowIndex = 1;
            EndColumnIndex = 1;

            Formula = string.Empty;
            NumberingCache = new SLNumberingCache();
        }

        internal string WorksheetName { get; set; }
        internal int StartRowIndex { get; set; }
        internal int StartColumnIndex { get; set; }
        internal int EndRowIndex { get; set; }
        internal int EndColumnIndex { get; set; }

        internal string Formula { get; set; }
        internal SLNumberingCache NumberingCache { get; set; }

        internal C.NumberReference ToNumberReference()
        {
            var nr = new C.NumberReference();
            nr.Formula = new C.Formula(Formula);
            nr.NumberingCache = NumberingCache.ToNumberingCache();

            return nr;
        }

        internal void RefreshFormula()
        {
            Formula = SLChartTool.GetChartReferenceFormula(WorksheetName, StartRowIndex, StartColumnIndex, EndRowIndex,
                EndColumnIndex);
        }

        internal SLNumberReference Clone()
        {
            var nr = new SLNumberReference();
            nr.WorksheetName = WorksheetName;
            nr.StartRowIndex = StartRowIndex;
            nr.StartColumnIndex = StartColumnIndex;
            nr.EndRowIndex = EndRowIndex;
            nr.EndColumnIndex = EndColumnIndex;
            nr.Formula = Formula;
            nr.NumberingCache = NumberingCache.Clone();

            return nr;
        }
    }
}