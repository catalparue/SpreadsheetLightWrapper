using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Ups.Toolkit.SpreadsheetLight.Core.misc;

namespace Ups.Toolkit.SpreadsheetLight.Core.table
{
    /// <summary>
    ///     Totals row function types.
    /// </summary>
    public enum SLTotalsRowFunctionValues
    {
        /// <summary>
        ///     Average
        /// </summary>
        Average = 0,

        /// <summary>
        ///     Count non-empty cells
        /// </summary>
        Count,

        /// <summary>
        ///     Count numbers
        /// </summary>
        CountNumbers,

        /// <summary>
        ///     Maximum
        /// </summary>
        Maximum,

        /// <summary>
        ///     Minimum
        /// </summary>
        Minimum,

        /// <summary>
        ///     Standard deviation
        /// </summary>
        StandardDeviation,

        /// <summary>
        ///     Sum
        /// </summary>
        Sum,

        /// <summary>
        ///     Variance
        /// </summary>
        Variance
    }

    /// <summary>
    ///     Encapsulates properties and methods for specifying tables. This simulates the
    ///     DocumentFormat.OpenXml.Spreadsheet.Table class.
    /// </summary>
    public class SLTable
    {
        internal bool HasSortState;

        internal bool HasTableStyleInfo;

        internal bool HasTableType;
        internal bool IsNewTable;

        internal string sDisplayName;
        private TableValues vTableType;

        internal SLTable()
        {
            SetAllNull();
        }

        internal string RelationshipID { get; set; }

        /// <summary>
        ///     Indicates if the table has auto-filter.
        /// </summary>
        public bool HasAutoFilter { get; set; }

        internal SLAutoFilter AutoFilter { get; set; }
        internal SLSortState SortState { get; set; }

        internal List<SLTableColumn> TableColumns { get; set; }
        internal HashSet<string> TableNames { get; set; }
        internal SLTableStyleInfo TableStyleInfo { get; set; }

        internal uint Id { get; set; }
        internal string Name { get; set; }

        /// <summary>
        ///     There should be no spaces in the given value.
        ///     Because display names of tables have to be unique across the entire spreadsheet,
        ///     this can only be checked when the table is actually inserted into the worksheet.
        ///     If the display name is duplicate, a new display name will be automatically assigned upon insertion.
        /// </summary>
        public string DisplayName
        {
            get { return sDisplayName; }
            set
            {
                sDisplayName = value;
                Name = sDisplayName;
            }
        }

        // The maximum length of this string should be 32,767 characters
        // We're not going to check this... TODO ?
        internal string Comment { get; set; }

        internal int StartRowIndex { get; set; }
        internal int StartColumnIndex { get; set; }
        internal int EndRowIndex { get; set; }
        internal int EndColumnIndex { get; set; }

        internal TableValues TableType
        {
            get { return vTableType; }
            set
            {
                vTableType = value;
                HasTableType = vTableType != TableValues.Worksheet ? true : false;
            }
        }

        internal uint HeaderRowCount { get; set; }
        internal bool? InsertRow { get; set; }
        internal bool? InsertRowShift { get; set; }

        internal uint TotalsRowCount { get; set; }
        internal bool? TotalsRowShown { get; set; }

        /// <summary>
        ///     Indicates if the table has a totals row.
        /// </summary>
        public bool HasTotalRow
        {
            get { return TotalsRowCount > 0 ? true : false; }
            set
            {
                if (value)
                {
                    // When inserting the table, the full table collision checks will be done.
                    if (TotalsRowCount == 0)
                    {
                        var iTotalsRowIndex = EndRowIndex + 1;
                        if (iTotalsRowIndex <= SLConstants.RowLimit)
                        {
                            EndRowIndex += 1;
                            TotalsRowCount = 1;
                            TotalsRowShown = true;
                        }
                    }
                }
                else
                {
                    if (TotalsRowCount > 0)
                    {
                        EndRowIndex -= 1;
                        // keep it at least one row deep
                        if (HeaderRowCount > 0)
                        {
                            if (EndRowIndex <= StartRowIndex) EndRowIndex = StartRowIndex + 1;
                        }
                        else
                        {
                            if (EndRowIndex < StartRowIndex) EndRowIndex = StartRowIndex;
                        }

                        TotalsRowCount = 0;
                        // no need to set TotalsRowShown false because it's a historic flag.
                        // If the totals row is *ever* shown, set it to true.
                    }
                    // else totals row count is already zero
                }
            }
        }

        internal bool? Published { get; set; }
        internal uint? HeaderRowFormatId { get; set; }
        internal uint? DataFormatId { get; set; }
        internal uint? TotalsRowFormatId { get; set; }
        internal uint? HeaderRowBorderFormatId { get; set; }
        internal uint? BorderFormatId { get; set; }
        internal uint? TotalsRowBorderFormatId { get; set; }
        internal string HeaderRowCellStyle { get; set; }
        internal string DataCellStyle { get; set; }
        internal string TotalsRowCellStyle { get; set; }
        internal uint? ConnectionId { get; set; }

        /// <summary>
        ///     Indicates if the table has banded rows.
        /// </summary>
        public bool HasBandedRows
        {
            get
            {
                // we'll default to true
                if (HasTableStyleInfo)
                    return TableStyleInfo.ShowRowStripes != null ? TableStyleInfo.ShowRowStripes.Value : true;
                return true;
            }
            set
            {
                TableStyleInfo.ShowRowStripes = value;
                HasTableStyleInfo = true;
            }
        }

        /// <summary>
        ///     Indicates if the table has banded columns.
        /// </summary>
        public bool HasBandedColumns
        {
            get
            {
                // we'll default to false
                if (HasTableStyleInfo)
                    return TableStyleInfo.ShowColumnStripes != null ? TableStyleInfo.ShowColumnStripes.Value : false;
                return false;
            }
            set
            {
                TableStyleInfo.ShowColumnStripes = value;
                HasTableStyleInfo = true;
            }
        }

        /// <summary>
        ///     Indicates if the table has special formatting for the first column.
        /// </summary>
        public bool HasFirstColumnStyled
        {
            get
            {
                // we'll default to false
                if (HasTableStyleInfo)
                    return TableStyleInfo.ShowFirstColumn != null ? TableStyleInfo.ShowFirstColumn.Value : false;
                return false;
            }
            set
            {
                TableStyleInfo.ShowFirstColumn = value;
                HasTableStyleInfo = true;
            }
        }

        /// <summary>
        ///     Indicates if the table has special formatting for the last column.
        /// </summary>
        public bool HasLastColumnStyled
        {
            get
            {
                // we'll default to false
                if (HasTableStyleInfo)
                    return TableStyleInfo.ShowLastColumn != null ? TableStyleInfo.ShowLastColumn.Value : false;
                return false;
            }
            set
            {
                TableStyleInfo.ShowLastColumn = value;
                HasTableStyleInfo = true;
            }
        }

        internal void SetAllNull()
        {
            IsNewTable = true;
            RelationshipID = string.Empty;

            AutoFilter = new SLAutoFilter();
            HasAutoFilter = false;
            SortState = new SLSortState();
            HasSortState = false;
            TableColumns = new List<SLTableColumn>();
            TableNames = new HashSet<string>();
            TableStyleInfo = new SLTableStyleInfo();
            HasTableStyleInfo = false;

            Id = 0;
            Name = null;
            sDisplayName = string.Empty;
            Comment = null;
            StartRowIndex = 1;
            StartColumnIndex = 1;
            EndRowIndex = 1;
            EndColumnIndex = 1;
            TableType = TableValues.Worksheet;
            HasTableType = false;
            HeaderRowCount = 1;
            InsertRow = null;
            InsertRowShift = null;
            TotalsRowCount = 0;
            TotalsRowShown = null;
            Published = null;
            HeaderRowFormatId = null;
            DataFormatId = null;
            TotalsRowFormatId = null;
            HeaderRowBorderFormatId = null;
            BorderFormatId = null;
            TotalsRowBorderFormatId = null;
            HeaderRowCellStyle = null;
            DataCellStyle = null;
            TotalsRowCellStyle = null;
            ConnectionId = null;
        }

        /// <summary>
        ///     Set the table style with a built-in style.
        /// </summary>
        /// <param name="TableStyle">A built-in table style.</param>
        public void SetTableStyle(SLTableStyleTypeValues TableStyle)
        {
            TableStyleInfo.SetTableStyle(TableStyle);
            HasTableStyleInfo = true;
        }

        /// <summary>
        ///     Remove the label text or function in the totals row.
        /// </summary>
        /// <param name="TableColumnIndex">
        ///     The table column index. For example, 1 for the 1st table column, 2 for the 2nd table
        ///     column and so on.
        /// </param>
        public void RemoveTotalRowLabelFunction(int TableColumnIndex)
        {
            --TableColumnIndex;
            if ((TableColumnIndex < 0) || (TableColumnIndex >= TableColumns.Count)) return;

            TableColumns[TableColumnIndex].TotalsRowLabel = null;
            TableColumns[TableColumnIndex].TotalsRowFunction = TotalsRowFunctionValues.None;
            TableColumns[TableColumnIndex].HasTotalsRowFunction = false;
        }

        /// <summary>
        ///     Set the label text in the totals row. Be sure to set <see cref="HasTotalRow" /> true first.
        /// </summary>
        /// <param name="TableColumnIndex">
        ///     The table column index. For example, 1 for the 1st table column, 2 for the 2nd table
        ///     column and so on.
        /// </param>
        /// <param name="Label">The label text.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool SetTotalRowLabel(int TableColumnIndex, string Label)
        {
            if (TotalsRowCount > 0)
            {
                --TableColumnIndex;
                if ((TableColumnIndex < 0) || (TableColumnIndex >= TableColumns.Count)) return false;

                TableColumns[TableColumnIndex].TotalsRowLabel = Label;
                TableColumns[TableColumnIndex].TotalsRowFunction = TotalsRowFunctionValues.None;
                TableColumns[TableColumnIndex].HasTotalsRowFunction = false;

                return true;
            }
            return false;
        }

        /// <summary>
        ///     Set the function in the totals row. Be sure to set <see cref="HasTotalRow" /> true first.
        /// </summary>
        /// <param name="TableColumnIndex">
        ///     The table column index. For example, 1 for the 1st table column, 2 for the 2nd table
        ///     column and so on.
        /// </param>
        /// <param name="TotalsRowFunction">The function type.</param>
        /// <returns>True if successful. False otherwise.</returns>
        public bool SetTotalRowFunction(int TableColumnIndex, SLTotalsRowFunctionValues TotalsRowFunction)
        {
            if (TotalsRowCount > 0)
            {
                --TableColumnIndex;
                if ((TableColumnIndex < 0) || (TableColumnIndex >= TableColumns.Count)) return false;

                TableColumns[TableColumnIndex].TotalsRowLabel = null;

                var iStartRowIndex = -1;
                var iEndRowIndex = -1;
                if (HeaderRowCount > 0) iStartRowIndex = StartRowIndex + 1;
                else iStartRowIndex = StartRowIndex;
                // not inclusive of the last totals row
                iEndRowIndex = EndRowIndex - 1;

                var iColumnIndex = StartColumnIndex + TableColumnIndex;

                switch (TotalsRowFunction)
                {
                    case SLTotalsRowFunctionValues.Average:
                        TableColumns[TableColumnIndex].TotalsRowFunction = TotalsRowFunctionValues.Average;
                        break;
                    case SLTotalsRowFunctionValues.Count:
                        TableColumns[TableColumnIndex].TotalsRowFunction = TotalsRowFunctionValues.Count;
                        break;
                    case SLTotalsRowFunctionValues.CountNumbers:
                        TableColumns[TableColumnIndex].TotalsRowFunction = TotalsRowFunctionValues.CountNumbers;
                        break;
                    case SLTotalsRowFunctionValues.Maximum:
                        TableColumns[TableColumnIndex].TotalsRowFunction = TotalsRowFunctionValues.Maximum;
                        break;
                    case SLTotalsRowFunctionValues.Minimum:
                        TableColumns[TableColumnIndex].TotalsRowFunction = TotalsRowFunctionValues.Minimum;
                        break;
                    case SLTotalsRowFunctionValues.StandardDeviation:
                        TableColumns[TableColumnIndex].TotalsRowFunction = TotalsRowFunctionValues.StandardDeviation;
                        break;
                    case SLTotalsRowFunctionValues.Sum:
                        TableColumns[TableColumnIndex].TotalsRowFunction = TotalsRowFunctionValues.Sum;
                        break;
                    case SLTotalsRowFunctionValues.Variance:
                        TableColumns[TableColumnIndex].TotalsRowFunction = TotalsRowFunctionValues.Variance;
                        break;
                }
                TableColumns[TableColumnIndex].HasTotalsRowFunction = true;

                return true;
            }
            return false;
        }

        /// <summary>
        ///     To sort data within the table. Note that the sorting is only done when the table is inserted into the worksheet.
        /// </summary>
        /// <param name="TableColumnIndex">
        ///     The table column index. For example, 1 for the 1st table column, 2 for the 2nd table
        ///     column and so on.
        /// </param>
        /// <param name="SortAscending">True to sort in ascending order. False to sort in descending order.</param>
        public void Sort(int TableColumnIndex, bool SortAscending)
        {
            --TableColumnIndex;
            if ((TableColumnIndex < 0) || (TableColumnIndex >= TableColumns.Count)) return;

            var iStartRowIndex = -1;
            var iEndRowIndex = -1;
            if (HeaderRowCount > 0) iStartRowIndex = StartRowIndex + 1;
            else iStartRowIndex = StartRowIndex;
            // not inclusive of the last totals row
            if (TotalsRowCount > 0) iEndRowIndex = EndRowIndex - 1;
            else iEndRowIndex = EndRowIndex;

            SortState = new SLSortState();
            SortState.StartRowIndex = iStartRowIndex;
            SortState.EndRowIndex = iEndRowIndex;
            SortState.StartColumnIndex = StartColumnIndex;
            SortState.EndColumnIndex = EndColumnIndex;

            var sc = new SLSortCondition();
            sc.StartRowIndex = iStartRowIndex;
            sc.StartColumnIndex = StartColumnIndex + TableColumnIndex;
            sc.EndRowIndex = iEndRowIndex;
            sc.EndColumnIndex = sc.StartColumnIndex;
            if (!SortAscending) sc.Descending = true;
            SortState.SortConditions.Add(sc);

            HasSortState = true;
        }

        internal void FromTable(Table t)
        {
            SetAllNull();

            if (t.AutoFilter != null)
            {
                AutoFilter.FromAutoFilter(t.AutoFilter);
                HasAutoFilter = true;
            }
            if (t.SortState != null)
            {
                SortState.FromSortState(t.SortState);
                HasSortState = true;
            }
            using (var oxr = OpenXmlReader.Create(t.TableColumns))
            {
                SLTableColumn tc;
                while (oxr.Read())
                    if (oxr.ElementType == typeof(TableColumn))
                    {
                        tc = new SLTableColumn();
                        tc.FromTableColumn((TableColumn) oxr.LoadCurrentElement());
                        TableColumns.Add(tc);
                    }
            }
            if (t.TableStyleInfo != null)
            {
                TableStyleInfo.FromTableStyleInfo(t.TableStyleInfo);
                HasTableStyleInfo = true;
            }

            Id = t.Id.Value;
            if (t.Name != null) Name = t.Name.Value;
            sDisplayName = t.DisplayName.Value;
            if (t.Comment != null) Comment = t.Comment.Value;

            var iStartRowIndex = 1;
            var iStartColumnIndex = 1;
            var iEndRowIndex = 1;
            var iEndColumnIndex = 1;
            var sRef = t.Reference.Value;
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

            if (t.TableType != null) TableType = t.TableType.Value;
            if ((t.HeaderRowCount != null) && (t.HeaderRowCount.Value != 1)) HeaderRowCount = t.HeaderRowCount.Value;
            if ((t.InsertRow != null) && t.InsertRow.Value) InsertRow = t.InsertRow.Value;
            if ((t.InsertRowShift != null) && t.InsertRowShift.Value) InsertRowShift = t.InsertRowShift.Value;
            if ((t.TotalsRowCount != null) && (t.TotalsRowCount.Value != 0)) TotalsRowCount = t.TotalsRowCount.Value;
            if ((t.TotalsRowShown != null) && !t.TotalsRowShown.Value) TotalsRowShown = t.TotalsRowShown.Value;
            if ((t.Published != null) && t.Published.Value) Published = t.Published.Value;
            if (t.HeaderRowFormatId != null) HeaderRowFormatId = t.HeaderRowFormatId.Value;
            if (t.DataFormatId != null) DataFormatId = t.DataFormatId.Value;
            if (t.TotalsRowFormatId != null) TotalsRowFormatId = t.TotalsRowFormatId.Value;
            if (t.HeaderRowBorderFormatId != null) HeaderRowBorderFormatId = t.HeaderRowBorderFormatId.Value;
            if (t.BorderFormatId != null) BorderFormatId = t.BorderFormatId.Value;
            if (t.TotalsRowBorderFormatId != null) TotalsRowBorderFormatId = t.TotalsRowBorderFormatId.Value;
            if (t.HeaderRowCellStyle != null) HeaderRowCellStyle = t.HeaderRowCellStyle.Value;
            if (t.DataCellStyle != null) DataCellStyle = t.DataCellStyle.Value;
            if (t.TotalsRowCellStyle != null) TotalsRowCellStyle = t.TotalsRowCellStyle.Value;
            if (t.ConnectionId != null) ConnectionId = t.ConnectionId.Value;
        }

        internal Table ToTable()
        {
            var t = new Table();
            if (HasAutoFilter) t.AutoFilter = AutoFilter.ToAutoFilter();
            if (HasSortState) t.SortState = SortState.ToSortState();

            t.TableColumns = new TableColumns {Count = (uint) TableColumns.Count};
            for (var i = 0; i < TableColumns.Count; ++i)
                t.TableColumns.Append(TableColumns[i].ToTableColumn());

            if (HasTableStyleInfo) t.TableStyleInfo = TableStyleInfo.ToTableStyleInfo();

            t.Id = Id;
            if (Name != null) t.Name = Name;
            t.DisplayName = DisplayName;
            if (Comment != null) t.Comment = Comment;

            if ((StartRowIndex == EndRowIndex) && (StartColumnIndex == EndColumnIndex))
                t.Reference = SLTool.ToCellReference(StartRowIndex, StartColumnIndex);
            else
                t.Reference = string.Format("{0}:{1}",
                    SLTool.ToCellReference(StartRowIndex, StartColumnIndex),
                    SLTool.ToCellReference(EndRowIndex, EndColumnIndex));

            if (HasTableType) t.TableType = TableType;
            if (HeaderRowCount != 1) t.HeaderRowCount = HeaderRowCount;
            if ((InsertRow != null) && InsertRow.Value) t.InsertRow = InsertRow.Value;
            if ((InsertRowShift != null) && InsertRowShift.Value) t.InsertRowShift = InsertRowShift.Value;
            if (TotalsRowCount != 0) t.TotalsRowCount = TotalsRowCount;
            if ((TotalsRowShown != null) && !TotalsRowShown.Value) t.TotalsRowShown = TotalsRowShown.Value;
            if ((Published != null) && Published.Value) t.Published = Published.Value;
            if (HeaderRowFormatId != null) t.HeaderRowFormatId = HeaderRowFormatId.Value;
            if (DataFormatId != null) t.DataFormatId = DataFormatId.Value;
            if (TotalsRowFormatId != null) t.TotalsRowFormatId = TotalsRowFormatId.Value;
            if (HeaderRowBorderFormatId != null) t.HeaderRowBorderFormatId = HeaderRowBorderFormatId.Value;
            if (BorderFormatId != null) t.BorderFormatId = BorderFormatId.Value;
            if (TotalsRowBorderFormatId != null) t.TotalsRowBorderFormatId = TotalsRowBorderFormatId.Value;
            if (HeaderRowCellStyle != null) t.HeaderRowCellStyle = HeaderRowCellStyle;
            if (DataCellStyle != null) t.DataCellStyle = DataCellStyle;
            if (TotalsRowCellStyle != null) t.TotalsRowCellStyle = TotalsRowCellStyle;
            if (ConnectionId != null) t.ConnectionId = ConnectionId.Value;

            return t;
        }

        internal SLTable Clone()
        {
            var t = new SLTable();
            t.IsNewTable = IsNewTable;
            t.RelationshipID = RelationshipID;
            t.HasAutoFilter = HasAutoFilter;
            t.AutoFilter = AutoFilter.Clone();
            t.HasSortState = HasSortState;
            t.SortState = SortState.Clone();

            t.TableColumns = new List<SLTableColumn>();
            for (var i = 0; i < TableColumns.Count; ++i)
                t.TableColumns.Add(TableColumns[i].Clone());

            t.TableNames = new HashSet<string>();
            foreach (var s in TableNames)
                t.TableNames.Add(s);

            t.HasTableStyleInfo = HasTableStyleInfo;
            t.TableStyleInfo = TableStyleInfo.Clone();

            t.Id = Id;
            t.Name = Name;
            t.sDisplayName = sDisplayName;
            t.Comment = Comment;
            t.StartRowIndex = StartRowIndex;
            t.StartColumnIndex = StartColumnIndex;
            t.EndRowIndex = EndRowIndex;
            t.EndColumnIndex = EndColumnIndex;

            t.HasTableType = HasTableType;
            t.vTableType = vTableType;

            t.HeaderRowCount = HeaderRowCount;
            t.InsertRow = InsertRow;
            t.InsertRowShift = InsertRowShift;
            t.TotalsRowCount = TotalsRowCount;
            t.TotalsRowShown = TotalsRowShown;

            t.Published = Published;
            t.HeaderRowFormatId = HeaderRowFormatId;
            t.DataFormatId = DataFormatId;
            t.TotalsRowFormatId = TotalsRowFormatId;
            t.HeaderRowBorderFormatId = HeaderRowBorderFormatId;
            t.BorderFormatId = BorderFormatId;
            t.TotalsRowBorderFormatId = TotalsRowBorderFormatId;
            t.HeaderRowCellStyle = HeaderRowCellStyle;
            t.DataCellStyle = DataCellStyle;
            t.TotalsRowCellStyle = TotalsRowCellStyle;
            t.ConnectionId = ConnectionId;

            return t;
        }
    }
}