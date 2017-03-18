using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.worksheet
{
    /// <summary>
    ///     Encapsulates properties and methods for specifying worksheet protection. This simulates the
    ///     DocumentFormat.OpenXml.Spreadsheet.SheetProtection class.
    /// </summary>
    public class SLSheetProtection
    {
        internal bool? bAllowAutoFilter;

        internal bool? bAllowDeleteColumns;

        internal bool? bAllowDeleteRows;

        // all the following properties take the negation of the default boolean value
        // of the corresponding attributes

        internal bool? bAllowEditObjects;

        internal bool? bAllowEditScenarios;

        internal bool? bAllowFormatCells;

        internal bool? bAllowFormatColumns;

        internal bool? bAllowFormatRows;

        internal bool? bAllowInsertColumns;

        internal bool? bAllowInsertHyperlinks;

        internal bool? bAllowInsertRows;

        internal bool? bAllowPivotTables;

        internal bool? bAllowSelectLockedCells;

        internal bool? bAllowSelectUnlockedCells;

        internal bool? bAllowSort;

        /// <summary>
        ///     Initializes an instance of SLSheetProtection.
        /// </summary>
        public SLSheetProtection()
        {
            SetAllNull();
        }

        internal string AlgorithmName { get; set; }
        internal string HashValue { get; set; }
        internal string SaltValue { get; set; }
        internal uint? SpinCount { get; set; }
        internal string Password { get; set; }

        internal bool? Sheet { get; set; }

        /// <summary>
        ///     Allow editing of objects even if sheet is protected.
        /// </summary>
        public bool AllowEditObjects
        {
            get { return bAllowEditObjects ?? true; }
            set { bAllowEditObjects = value; }
        }

        /// <summary>
        ///     Allow editing of scenarios even if sheet is protected.
        /// </summary>
        public bool AllowEditScenarios
        {
            get { return bAllowEditScenarios ?? true; }
            set { bAllowEditScenarios = value; }
        }

        /// <summary>
        ///     Allow formatting of cells even if sheet is protected.
        /// </summary>
        public bool AllowFormatCells
        {
            get { return bAllowFormatCells ?? false; }
            set { bAllowFormatCells = value; }
        }

        /// <summary>
        ///     Allow formatting of columns even if sheet is protected.
        /// </summary>
        public bool AllowFormatColumns
        {
            get { return bAllowFormatColumns ?? false; }
            set { bAllowFormatColumns = value; }
        }

        /// <summary>
        ///     Allow formatting of rows even if sheet is protected.
        /// </summary>
        public bool AllowFormatRows
        {
            get { return bAllowFormatRows ?? false; }
            set { bAllowFormatRows = value; }
        }

        /// <summary>
        ///     Allow insertion of columns even if sheet is protected.
        /// </summary>
        public bool AllowInsertColumns
        {
            get { return bAllowInsertColumns ?? false; }
            set { bAllowInsertColumns = value; }
        }

        /// <summary>
        ///     Allow insertion of rows even if sheet is protected.
        /// </summary>
        public bool AllowInsertRows
        {
            get { return bAllowInsertRows ?? false; }
            set { bAllowInsertRows = value; }
        }

        /// <summary>
        ///     Allow insertion of hyperlinks even if sheet is protected.
        /// </summary>
        public bool AllowInsertHyperlinks
        {
            get { return bAllowInsertHyperlinks ?? false; }
            set { bAllowInsertHyperlinks = value; }
        }

        /// <summary>
        ///     Allow deletion of columns even if sheet is protected.
        /// </summary>
        public bool AllowDeleteColumns
        {
            get { return bAllowDeleteColumns ?? false; }
            set { bAllowDeleteColumns = value; }
        }

        /// <summary>
        ///     Allow deletion of rows even if sheet is protected.
        /// </summary>
        public bool AllowDeleteRows
        {
            get { return bAllowDeleteRows ?? false; }
            set { bAllowDeleteRows = value; }
        }

        /// <summary>
        ///     Allow selection of locked cells even if sheet is protected.
        /// </summary>
        public bool AllowSelectLockedCells
        {
            get { return bAllowSelectLockedCells ?? true; }
            set { bAllowSelectLockedCells = value; }
        }

        /// <summary>
        ///     Allow sorting even if sheet is protected.
        /// </summary>
        public bool AllowSort
        {
            get { return bAllowSort ?? false; }
            set { bAllowSort = value; }
        }

        /// <summary>
        ///     Allow use of autofilters even if sheet is protected.
        /// </summary>
        public bool AllowAutoFilter
        {
            get { return bAllowAutoFilter ?? false; }
            set { bAllowAutoFilter = value; }
        }

        /// <summary>
        ///     Allow use of pivot tables even if sheet is protected.
        /// </summary>
        public bool AllowPivotTables
        {
            get { return bAllowPivotTables ?? false; }
            set { bAllowPivotTables = value; }
        }

        /// <summary>
        ///     Allow selection of unlocked cells even if sheet is protected.
        /// </summary>
        public bool AllowSelectUnlockedCells
        {
            get { return bAllowSelectUnlockedCells ?? true; }
            set { bAllowSelectUnlockedCells = value; }
        }

        internal void SetAllNull()
        {
            AlgorithmName = null;
            HashValue = null;
            SaltValue = null;
            SpinCount = null;
            Password = null;
            Sheet = null;
            bAllowEditObjects = null;
            bAllowEditScenarios = null;
            bAllowFormatCells = null;
            bAllowFormatColumns = null;
            bAllowFormatRows = null;
            bAllowInsertColumns = null;
            bAllowInsertRows = null;
            bAllowInsertHyperlinks = null;
            bAllowDeleteColumns = null;
            bAllowDeleteRows = null;
            bAllowSelectLockedCells = null;
            bAllowSort = null;
            bAllowAutoFilter = null;
            bAllowPivotTables = null;
            bAllowSelectUnlockedCells = null;
        }

        internal void FromSheetProtection(SheetProtection sp)
        {
            SetAllNull();
            if (sp.AlgorithmName != null) AlgorithmName = sp.AlgorithmName.Value;
            if (sp.HashValue != null) HashValue = sp.HashValue.Value;
            if (sp.SaltValue != null) SaltValue = sp.SaltValue.Value;
            if (sp.SpinCount != null) SpinCount = sp.SpinCount.Value;
            if (sp.Password != null) Password = sp.Password.Value;
            if (sp.Sheet != null) Sheet = sp.Sheet.Value;

            if (sp.Objects != null) AllowEditObjects = !sp.Objects.Value;
            if (sp.Scenarios != null) AllowEditScenarios = !sp.Scenarios.Value;
            if (sp.FormatCells != null) AllowFormatCells = !sp.FormatCells.Value;
            if (sp.FormatColumns != null) AllowFormatColumns = !sp.FormatColumns.Value;
            if (sp.FormatRows != null) AllowFormatRows = !sp.FormatRows.Value;
            if (sp.InsertColumns != null) AllowInsertColumns = !sp.InsertColumns.Value;
            if (sp.InsertRows != null) AllowInsertRows = !sp.InsertRows.Value;
            if (sp.InsertHyperlinks != null) AllowInsertHyperlinks = !sp.InsertHyperlinks.Value;
            if (sp.DeleteColumns != null) AllowDeleteColumns = !sp.DeleteColumns.Value;
            if (sp.DeleteRows != null) AllowDeleteRows = !sp.DeleteRows.Value;
            if (sp.SelectLockedCells != null) AllowSelectLockedCells = !sp.SelectLockedCells.Value;
            if (sp.Sort != null) AllowSort = !sp.Sort.Value;
            if (sp.AutoFilter != null) AllowAutoFilter = !sp.AutoFilter.Value;
            if (sp.PivotTables != null) AllowPivotTables = !sp.PivotTables.Value;
            if (sp.SelectUnlockedCells != null) AllowSelectUnlockedCells = !sp.SelectUnlockedCells.Value;
        }

        internal SheetProtection ToSheetProtection()
        {
            var sp = new SheetProtection();
            if (AlgorithmName != null) sp.AlgorithmName = AlgorithmName;
            if (HashValue != null) sp.HashValue = HashValue;
            if (SaltValue != null) sp.SaltValue = SaltValue;
            if (SpinCount != null) sp.SpinCount = SpinCount.Value;
            if (Password != null) sp.Password = Password;
            if ((Sheet != null) && Sheet.Value) sp.Sheet = Sheet.Value;

            if (!AllowEditObjects) sp.Objects = !AllowEditObjects;
            if (!AllowEditScenarios) sp.Scenarios = !AllowEditScenarios;
            if (!AllowFormatCells != true) sp.FormatCells = !AllowFormatCells;
            if (!AllowFormatColumns != true) sp.FormatColumns = !AllowFormatColumns;
            if (!AllowFormatRows != true) sp.FormatRows = !AllowFormatRows;
            if (!AllowInsertColumns != true) sp.InsertColumns = !AllowInsertColumns;
            if (!AllowInsertRows != true) sp.InsertRows = !AllowInsertRows;
            if (!AllowInsertHyperlinks != true) sp.InsertHyperlinks = !AllowInsertHyperlinks;
            if (!AllowDeleteColumns != true) sp.DeleteColumns = !AllowDeleteColumns;
            if (!AllowDeleteRows != true) sp.DeleteRows = !AllowDeleteRows;
            if (!AllowSelectLockedCells) sp.SelectLockedCells = !AllowSelectLockedCells;
            if (!AllowSort != true) sp.Sort = !AllowSort;
            if (!AllowAutoFilter != true) sp.AutoFilter = !AllowAutoFilter;
            if (!AllowPivotTables != true) sp.PivotTables = !AllowPivotTables;
            if (!AllowSelectUnlockedCells) sp.SelectUnlockedCells = !AllowSelectUnlockedCells;

            return sp;
        }

        internal SLSheetProtection Clone()
        {
            var sp = new SLSheetProtection();
            sp.AlgorithmName = AlgorithmName;
            sp.HashValue = HashValue;
            sp.SaltValue = SaltValue;
            sp.SpinCount = SpinCount;
            sp.Password = Password;
            sp.Sheet = Sheet;
            sp.bAllowEditObjects = bAllowEditObjects;
            sp.bAllowEditScenarios = bAllowEditScenarios;
            sp.bAllowFormatCells = bAllowFormatCells;
            sp.bAllowFormatColumns = bAllowFormatColumns;
            sp.bAllowFormatRows = bAllowFormatRows;
            sp.bAllowInsertColumns = bAllowInsertColumns;
            sp.bAllowInsertRows = bAllowInsertRows;
            sp.bAllowInsertHyperlinks = bAllowInsertHyperlinks;
            sp.bAllowDeleteColumns = bAllowDeleteColumns;
            sp.bAllowDeleteRows = bAllowDeleteRows;
            sp.bAllowSelectLockedCells = bAllowSelectLockedCells;
            sp.bAllowSort = bAllowSort;
            sp.bAllowAutoFilter = bAllowAutoFilter;
            sp.bAllowPivotTables = bAllowPivotTables;
            sp.bAllowSelectUnlockedCells = bAllowSelectUnlockedCells;

            return sp;
        }
    }
}