using DocumentFormat.OpenXml.Spreadsheet;
using Ups.Toolkit.SpreadsheetLight.Core.misc;
using Ups.Toolkit.SpreadsheetLight.Core.worksheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.pivottable
{
    internal class SLLocation
    {
        internal SLLocation()
        {
            SetAllNull();
        }

        internal SLCellPointRange Reference { get; set; }
        internal uint FirstHeaderRow { get; set; }
        internal uint FirstDataRow { get; set; }
        internal uint FirstDataColumn { get; set; }
        internal uint RowPageCount { get; set; }
        internal uint ColumnsPerPage { get; set; }

        private void SetAllNull()
        {
            Reference = new SLCellPointRange(1, 1, 1, 1);
            FirstHeaderRow = 1;
            FirstDataRow = 1;
            FirstDataColumn = 1;
            RowPageCount = 0;
            ColumnsPerPage = 0;
        }

        internal void FromLocation(Location loc)
        {
            SetAllNull();

            if (loc.Reference != null) Reference = SLTool.TranslateReferenceToCellPointRange(loc.Reference.Value);
            if (loc.FirstHeaderRow != null) FirstHeaderRow = loc.FirstHeaderRow.Value;
            if (loc.FirstDataRow != null) FirstDataRow = loc.FirstDataRow.Value;
            if (loc.FirstDataColumn != null) FirstDataColumn = loc.FirstDataColumn.Value;
            if (loc.RowPageCount != null) RowPageCount = loc.RowPageCount.Value;
            if (loc.ColumnsPerPage != null) ColumnsPerPage = loc.ColumnsPerPage.Value;
        }

        internal Location ToLocation()
        {
            var loc = new Location();
            if ((Reference.StartRowIndex == Reference.EndRowIndex)
                && (Reference.StartColumnIndex == Reference.EndColumnIndex))
                loc.Reference = SLTool.ToCellReference(Reference.StartRowIndex, Reference.StartColumnIndex);
            else
                loc.Reference = SLTool.ToCellRange(Reference.StartRowIndex, Reference.StartColumnIndex,
                    Reference.EndRowIndex, Reference.EndColumnIndex);

            loc.FirstHeaderRow = FirstHeaderRow;
            loc.FirstDataRow = FirstDataRow;
            loc.FirstDataColumn = FirstDataColumn;
            if (RowPageCount != 0) loc.RowPageCount = RowPageCount;
            if (ColumnsPerPage != 0) loc.ColumnsPerPage = ColumnsPerPage;

            return loc;
        }

        internal SLLocation Clone()
        {
            var loc = new SLLocation();
            loc.Reference = new SLCellPointRange(Reference.StartRowIndex, Reference.StartColumnIndex,
                Reference.EndRowIndex, Reference.EndColumnIndex);
            loc.FirstHeaderRow = FirstHeaderRow;
            loc.FirstDataRow = FirstDataRow;
            loc.FirstDataColumn = FirstDataColumn;
            loc.RowPageCount = RowPageCount;
            loc.ColumnsPerPage = ColumnsPerPage;

            return loc;
        }
    }
}