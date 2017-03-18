using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLightWrapper.Core.table;

namespace SpreadsheetLightWrapper.Core.misc
{
    internal class SLAutoFilter
    {
        internal bool HasSortState;

        internal SLAutoFilter()
        {
            SetAllNull();
        }

        internal List<SLFilterColumn> FilterColumns { get; set; }
        internal SLSortState SortState { get; set; }

        internal int StartRowIndex { get; set; }
        internal int StartColumnIndex { get; set; }
        internal int EndRowIndex { get; set; }
        internal int EndColumnIndex { get; set; }

        private void SetAllNull()
        {
            FilterColumns = new List<SLFilterColumn>();
            SortState = new SLSortState();
            HasSortState = false;
            StartRowIndex = 1;
            StartColumnIndex = 1;
            EndRowIndex = 1;
            EndColumnIndex = 1;
        }

        internal void FromAutoFilter(AutoFilter af)
        {
            SetAllNull();

            var iStartRowIndex = 1;
            var iStartColumnIndex = 1;
            var iEndRowIndex = 1;
            var iEndColumnIndex = 1;
            var sRef = af.Reference.Value;
            if (sRef.IndexOf(":") > 0)
            {
                if (SLTool.FormatCellReferenceRangeToRowColumnIndex(sRef, out iStartRowIndex, out iStartColumnIndex,
                    out iEndRowIndex, out iEndColumnIndex))
                {
                    StartRowIndex = iStartRowIndex;
                    StartColumnIndex = iStartColumnIndex;
                    EndRowIndex = iEndRowIndex;
                    EndColumnIndex = iEndColumnIndex;
                }
            }
            else
            {
                if (SLTool.FormatCellReferenceToRowColumnIndex(sRef, out iStartRowIndex, out iStartColumnIndex))
                {
                    StartRowIndex = iStartRowIndex;
                    StartColumnIndex = iStartColumnIndex;
                    EndRowIndex = iStartRowIndex;
                    EndColumnIndex = iStartColumnIndex;
                }
            }

            if (af.HasChildren)
            {
                SLFilterColumn fc;
                using (var oxr = OpenXmlReader.Create(af))
                {
                    while (oxr.Read())
                        if (oxr.ElementType == typeof(FilterColumn))
                        {
                            fc = new SLFilterColumn();
                            fc.FromFilterColumn((FilterColumn) oxr.LoadCurrentElement());
                            FilterColumns.Add(fc);
                        }
                        else if (oxr.ElementType == typeof(SortState))
                        {
                            SortState = new SLSortState();
                            SortState.FromSortState((SortState) oxr.LoadCurrentElement());
                            HasSortState = true;
                        }
                }
            }
        }

        internal AutoFilter ToAutoFilter()
        {
            var af = new AutoFilter();

            if ((StartRowIndex == EndRowIndex) && (StartColumnIndex == EndColumnIndex))
                af.Reference = SLTool.ToCellReference(StartRowIndex, StartColumnIndex);
            else
                af.Reference = string.Format("{0}:{1}",
                    SLTool.ToCellReference(StartRowIndex, StartColumnIndex),
                    SLTool.ToCellReference(EndRowIndex, EndColumnIndex));

            foreach (var fc in FilterColumns)
                af.Append(fc.ToFilterColumn());

            if (HasSortState) af.Append(SortState.ToSortState());

            return af;
        }

        internal SLAutoFilter Clone()
        {
            var af = new SLAutoFilter();
            af.FilterColumns = new List<SLFilterColumn>();
            for (var i = 0; i < FilterColumns.Count; ++i)
                af.FilterColumns.Add(FilterColumns[i].Clone());

            af.HasSortState = HasSortState;
            af.SortState = SortState.Clone();

            af.StartRowIndex = StartRowIndex;
            af.StartColumnIndex = StartColumnIndex;
            af.EndRowIndex = EndRowIndex;
            af.EndColumnIndex = EndColumnIndex;

            return af;
        }
    }
}