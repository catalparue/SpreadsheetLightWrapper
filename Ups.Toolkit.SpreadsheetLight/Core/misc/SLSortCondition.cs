using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.misc
{
    internal class SLSortCondition
    {
        internal bool HasIconSet;

        internal bool HasSortBy;
        private IconSetValues vIconSet;
        private SortByValues vSortBy;

        internal SLSortCondition()
        {
            SetAllNull();
        }

        internal bool? Descending { get; set; }

        internal SortByValues SortBy
        {
            get { return vSortBy; }
            set
            {
                vSortBy = value;
                HasSortBy = vSortBy != SortByValues.Value ? true : false;
            }
        }

        internal int StartRowIndex { get; set; }
        internal int StartColumnIndex { get; set; }
        internal int EndRowIndex { get; set; }
        internal int EndColumnIndex { get; set; }

        internal string CustomList { get; set; }
        internal uint? FormatId { get; set; }

        internal IconSetValues IconSet
        {
            get { return vIconSet; }
            set
            {
                vIconSet = value;
                HasIconSet = vIconSet != IconSetValues.ThreeArrows ? true : false;
            }
        }

        internal uint? IconId { get; set; }

        private void SetAllNull()
        {
            Descending = null;
            vSortBy = SortByValues.Value;
            HasSortBy = false;

            StartRowIndex = 1;
            StartColumnIndex = 1;
            EndRowIndex = 1;
            EndColumnIndex = 1;

            CustomList = null;
            FormatId = null;

            vIconSet = IconSetValues.ThreeArrows;
            HasIconSet = false;

            IconId = null;
        }

        internal void FromSortCondition(SortCondition sc)
        {
            SetAllNull();

            if ((sc.Descending != null) && sc.Descending.Value) Descending = sc.Descending.Value;
            if (sc.SortBy != null) SortBy = sc.SortBy.Value;

            var iStartRowIndex = 1;
            var iStartColumnIndex = 1;
            var iEndRowIndex = 1;
            var iEndColumnIndex = 1;
            var sRef = sc.Reference.Value;
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

            if (sc.CustomList != null) CustomList = sc.CustomList.Value;
            if (sc.FormatId != null) FormatId = sc.FormatId.Value;
            if (sc.IconSet != null) IconSet = sc.IconSet.Value;
            if (sc.IconId != null) IconId = sc.IconId.Value;
        }

        internal SortCondition ToSortCondition()
        {
            var sc = new SortCondition();
            if (Descending != null) sc.Descending = Descending.Value;
            if (HasSortBy) sc.SortBy = SortBy;

            if ((StartRowIndex == EndRowIndex) && (StartColumnIndex == EndColumnIndex))
                sc.Reference = SLTool.ToCellReference(StartRowIndex, StartColumnIndex);
            else
                sc.Reference = string.Format("{0}:{1}",
                    SLTool.ToCellReference(StartRowIndex, StartColumnIndex),
                    SLTool.ToCellReference(EndRowIndex, EndColumnIndex));

            if (CustomList != null) sc.CustomList = CustomList;
            if (FormatId != null) sc.FormatId = FormatId;
            if (HasIconSet) sc.IconSet = IconSet;
            if (IconId != null) sc.IconId = IconId.Value;

            return sc;
        }

        internal SLSortCondition Clone()
        {
            var sc = new SLSortCondition();
            sc.Descending = Descending;
            sc.HasSortBy = HasSortBy;
            sc.vSortBy = vSortBy;
            sc.StartRowIndex = StartRowIndex;
            sc.StartColumnIndex = StartColumnIndex;
            sc.EndRowIndex = EndRowIndex;
            sc.EndColumnIndex = EndColumnIndex;
            sc.CustomList = CustomList;
            sc.FormatId = FormatId;
            sc.HasIconSet = HasIconSet;
            sc.vIconSet = vIconSet;
            sc.IconId = IconId;

            return sc;
        }
    }
}