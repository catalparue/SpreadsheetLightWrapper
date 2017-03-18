using System.Collections.Generic;

namespace Ups.Toolkit.SpreadsheetLight.Core.workbook
{
    internal enum SLSheetType
    {
        Unknown = 0,
        Worksheet,
        Chartsheet,

        /// <summary>
        ///     I hope I don't have to support this... This complements macro-enabled Excel files?
        /// </summary>
        DialogSheet,

        /// <summary>
        ///     In future? Is this for macro-enabled Excel files?
        /// </summary>
        Macrosheet
    }

    internal class SLWorkbook
    {
        internal SLWorkbook()
        {
            WorkbookProperties = new SLWorkbookProperties();
            WorkbookViews = new List<SLWorkbookView>();
            Sheets = new List<SLSheet>();
            DefinedNames = new List<SLDefinedName>();
            CalculationCells = new List<SLCalculationCell>();

            PossibleTableId = 1;
            TableIds = new HashSet<uint>();
            TableNames = new HashSet<string>();

            PossiblePivotTableCacheId = 1;
            PivotTableNames = new HashSet<string>();
            PivotTableCacheIds = new HashSet<uint>();
        }

        internal SLWorkbookProperties WorkbookProperties { get; set; }
        internal List<SLWorkbookView> WorkbookViews { get; set; }
        internal List<SLSheet> Sheets { get; set; }
        internal List<SLDefinedName> DefinedNames { get; set; }
        internal List<SLCalculationCell> CalculationCells { get; set; }

        internal uint PossibleTableId { get; set; }
        internal HashSet<uint> TableIds { get; set; }
        internal HashSet<string> TableNames { get; set; }

        // Pivot tables can have the same names as normal tables.

        internal uint PossiblePivotTableCacheId { get; set; }
        internal HashSet<string> PivotTableNames { get; set; }
        internal HashSet<uint> PivotTableCacheIds { get; set; }

        internal void RefreshPossibleTableId()
        {
            PossibleTableId = 1;
            // possible infinite loop, but how many tables do you have anyway?
            while (TableIds.Contains(PossibleTableId))
                ++PossibleTableId;
        }

        /// <summary>
        ///     Call RefreshPossibleTableId() first!
        /// </summary>
        /// <returns></returns>
        internal string GetNextPossibleTableName()
        {
            var i = PossibleTableId;
            var sName = string.Format("Table{0}", i);
            // possible infinite loop, but how many tables do you have anyway?
            while (TableNames.Contains(sName))
            {
                ++i;
                sName = string.Format("Table{0}", i);
            }
            return sName;
        }

        internal bool HasTableName(string TableName)
        {
            return TableNames.Contains(TableName);
        }

        internal void RefreshPossiblePivotTableCacheId()
        {
            PossiblePivotTableCacheId = 1;
            // possible infinite loop, but how many pivot tables do you have anyway?
            while (PivotTableCacheIds.Contains(PossiblePivotTableCacheId))
                ++PossiblePivotTableCacheId;
        }

        /// <summary>
        ///     Call RefreshPossiblePivotTableId() first!
        /// </summary>
        /// <returns></returns>
        internal string GetNextPossiblePivotTableName()
        {
            //uint i = PossiblePivotTableCacheId;

            // The cache ID doesn't have to coincide with the pivot table name.
            // In fact, Excel usually has different numbers for them. And I don't know how
            // Excel comes up with the cache IDs.
            var i = (uint) PivotTableNames.Count + 1;
            var sName = string.Format("PivotTable{0}", i);
            // possible infinite loop, but how many pivot tables do you have anyway?
            while (PivotTableNames.Contains(sName))
            {
                ++i;
                sName = string.Format("PivotTable{0}", i);
            }
            return sName;
        }

        internal bool HasPivotTableName(string PivotTableName)
        {
            return PivotTableNames.Contains(PivotTableName);
        }

        /// <summary>
        ///     Adds a calculation cell if it doesn't already exist
        /// </summary>
        /// <param name="cc"></param>
        internal void AddCalculationCell(SLCalculationCell cc)
        {
            var bFound = false;
            foreach (var calc in CalculationCells)
                if ((calc.SheetId == cc.SheetId) && (calc.RowIndex == cc.RowIndex) &&
                    (calc.ColumnIndex == cc.ColumnIndex))
                {
                    bFound = true;
                    break;
                }

            if (!bFound) CalculationCells.Add(cc.Clone());
        }
    }
}