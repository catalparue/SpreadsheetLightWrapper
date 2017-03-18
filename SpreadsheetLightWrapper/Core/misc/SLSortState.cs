using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.misc
{
    internal class SLSortState
    {
        internal bool HasSortMethod;
        private SortMethodValues vSortMethod;

        internal SLSortState()
        {
            SetAllNull();
        }

        internal List<SLSortCondition> SortConditions { get; set; }
        internal bool? ColumnSort { get; set; }
        internal bool? CaseSensitive { get; set; }

        internal SortMethodValues SortMethod
        {
            get { return vSortMethod; }
            set
            {
                vSortMethod = value;
                HasSortMethod = vSortMethod != SortMethodValues.None ? true : false;
            }
        }

        internal int StartRowIndex { get; set; }
        internal int StartColumnIndex { get; set; }
        internal int EndRowIndex { get; set; }
        internal int EndColumnIndex { get; set; }

        private void SetAllNull()
        {
            SortConditions = new List<SLSortCondition>();
            ColumnSort = null;
            CaseSensitive = null;

            vSortMethod = SortMethodValues.None;
            HasSortMethod = false;

            StartRowIndex = 1;
            StartColumnIndex = 1;
            EndRowIndex = 1;
            EndColumnIndex = 1;
        }

        internal void FromSortState(SortState ss)
        {
            SetAllNull();

            if ((ss.ColumnSort != null) && ss.ColumnSort.Value) ColumnSort = ss.ColumnSort.Value;
            if ((ss.CaseSensitive != null) && ss.CaseSensitive.Value) CaseSensitive = ss.CaseSensitive.Value;
            if (ss.SortMethod != null) SortMethod = ss.SortMethod.Value;

            var iStartRowIndex = 1;
            var iStartColumnIndex = 1;
            var iEndRowIndex = 1;
            var iEndColumnIndex = 1;
            var sRef = ss.Reference.Value;
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

            if (ss.HasChildren)
            {
                SLSortCondition sc;
                using (var oxr = OpenXmlReader.Create(ss))
                {
                    while (oxr.Read())
                        if (oxr.ElementType == typeof(SortCondition))
                        {
                            sc = new SLSortCondition();
                            sc.FromSortCondition((SortCondition) oxr.LoadCurrentElement());
                            // limit of 64 from Open XML specs
                            if (SortConditions.Count < 64) SortConditions.Add(sc);
                        }
                }
            }
        }

        internal SortState ToSortState()
        {
            var ss = new SortState();
            if ((ColumnSort != null) && ColumnSort.Value) ss.ColumnSort = ColumnSort.Value;
            if ((CaseSensitive != null) && CaseSensitive.Value) ss.CaseSensitive = CaseSensitive.Value;
            if (HasSortMethod) ss.SortMethod = SortMethod;

            if ((StartRowIndex == EndRowIndex) && (StartColumnIndex == EndColumnIndex))
                ss.Reference = SLTool.ToCellReference(StartRowIndex, StartColumnIndex);
            else
                ss.Reference = string.Format("{0}:{1}",
                    SLTool.ToCellReference(StartRowIndex, StartColumnIndex),
                    SLTool.ToCellReference(EndRowIndex, EndColumnIndex));

            if (SortConditions.Count > 0)
                for (var i = 0; i < SortConditions.Count; ++i)
                    ss.Append(SortConditions[i].ToSortCondition());

            return ss;
        }

        internal SLSortState Clone()
        {
            var ss = new SLSortState();
            ss.SortConditions = new List<SLSortCondition>();
            for (var i = 0; i < SortConditions.Count; ++i)
                ss.SortConditions.Add(SortConditions[i].Clone());

            ss.ColumnSort = ColumnSort;
            ss.CaseSensitive = CaseSensitive;

            ss.HasSortMethod = HasSortMethod;
            ss.vSortMethod = vSortMethod;

            ss.StartRowIndex = StartRowIndex;
            ss.StartColumnIndex = StartColumnIndex;
            ss.EndRowIndex = EndRowIndex;
            ss.EndColumnIndex = EndColumnIndex;

            return ss;
        }
    }
}