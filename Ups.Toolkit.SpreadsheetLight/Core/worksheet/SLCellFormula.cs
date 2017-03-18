using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.worksheet
{
    /// <summary>
    ///     This is for information purposes only! This simulates the DocumentFormat.OpenXml.Spreadsheet.CellFormula class.
    /// </summary>
    public class SLCellFormula
    {
        internal SLCellFormula()
        {
            SetAllNull();
        }

        // We're not going to do preserving space. Don't know the full behaviour for
        // excessively spaced formula text...

        /// <summary>
        ///     The formula text.
        /// </summary>
        public string FormulaText { get; set; }

        /// <summary>
        ///     The type of formula.
        /// </summary>
        public CellFormulaValues FormulaType { get; set; }

        /// <summary>
        ///     If true, then formula is an array formula and the entire array is calculated in full.
        ///     If false, then individual cells are calculated as needed.
        /// </summary>
        public bool AlwaysCalculateArray { get; set; }

        /// <summary>
        ///     Range of cells where the formula is applied.
        /// </summary>
        public string Reference { get; set; }

        /// <summary>
        ///     True for 2-dimensional data table. False otherwise.
        /// </summary>
        public bool DataTable2D { get; set; }

        /// <summary>
        ///     If true, then 1-dimensional data table is a row. Otherwise it's a column.
        /// </summary>
        public bool DataTableRow { get; set; }

        /// <summary>
        ///     Whether the first input cell for data table is deleted.
        /// </summary>
        public bool Input1Deleted { get; set; }

        /// <summary>
        ///     Whether the second input cell for data table is deleted.
        /// </summary>
        public bool Input2Deleted { get; set; }

        /// <summary>
        ///     First input cell for data table.
        /// </summary>
        public string R1 { get; set; }

        /// <summary>
        ///     Second input cell for data table.
        /// </summary>
        public string R2 { get; set; }

        /// <summary>
        ///     Indicates whether this formula needs to be recalculated.
        /// </summary>
        public bool CalculateCell { get; set; }

        /// <summary>
        ///     Shared formula index.
        /// </summary>
        public uint? SharedIndex { get; set; }

        /// <summary>
        ///     Specifies that this formula assigns a value to a name.
        /// </summary>
        public bool Bx { get; set; }

        internal void SetAllNull()
        {
            FormulaText = string.Empty;

            FormulaType = CellFormulaValues.Normal;
            AlwaysCalculateArray = false;
            Reference = "";
            DataTable2D = false;
            DataTableRow = false;
            Input1Deleted = false;
            Input2Deleted = false;
            R1 = "";
            R2 = "";
            CalculateCell = false;
            SharedIndex = null;
            Bx = false;
        }

        internal void FromCellFormula(CellFormula cf)
        {
            SetAllNull();

            FormulaText = cf.Text;
            if (cf.FormulaType != null) FormulaType = cf.FormulaType.Value;
            if (cf.AlwaysCalculateArray != null) AlwaysCalculateArray = cf.AlwaysCalculateArray.Value;
            if (cf.Reference != null) Reference = cf.Reference.Value;
            if (cf.DataTable2D != null) DataTable2D = cf.DataTable2D.Value;
            if (cf.DataTableRow != null) DataTableRow = cf.DataTableRow.Value;
            if (cf.Input1Deleted != null) Input1Deleted = cf.Input1Deleted.Value;
            if (cf.Input2Deleted != null) Input2Deleted = cf.Input2Deleted.Value;
            if (cf.R1 != null) R1 = cf.R1.Value;
            if (cf.R2 != null) R2 = cf.R2.Value;
            if (cf.CalculateCell != null) CalculateCell = cf.CalculateCell.Value;
            if (cf.SharedIndex != null) SharedIndex = cf.SharedIndex.Value;
            if (cf.Bx != null) Bx = cf.Bx.Value;
        }

        internal CellFormula ToCellFormula()
        {
            var cf = new CellFormula();
            cf.Text = FormulaText;

            if (FormulaType != CellFormulaValues.Normal) cf.FormulaType = FormulaType;
            if (AlwaysCalculateArray) cf.AlwaysCalculateArray = AlwaysCalculateArray;
            if (Reference.Length > 0) cf.Reference = Reference;
            if (DataTable2D) cf.DataTable2D = DataTable2D;
            if (DataTableRow) cf.DataTableRow = DataTableRow;
            if (Input1Deleted) cf.Input1Deleted = Input1Deleted;
            if (Input2Deleted) cf.Input2Deleted = Input2Deleted;
            if (R1.Length > 0) cf.R1 = R1;
            if (R2.Length > 0) cf.R2 = R2;
            if (CalculateCell) cf.CalculateCell = CalculateCell;
            if (SharedIndex != null) cf.SharedIndex = SharedIndex.Value;
            if (Bx) cf.Bx = Bx;

            return cf;
        }

        internal static string GetFormulaTypeAttribute(CellFormulaValues cfv)
        {
            var result = "normal";
            switch (cfv)
            {
                case CellFormulaValues.Normal:
                    result = "normal";
                    break;
                case CellFormulaValues.Array:
                    result = "array";
                    break;
                case CellFormulaValues.DataTable:
                    result = "dataTable";
                    break;
                case CellFormulaValues.Shared:
                    result = "shared";
                    break;
            }

            return result;
        }

        internal SLCellFormula Clone()
        {
            var cf = new SLCellFormula();
            cf.FormulaText = FormulaText;
            cf.FormulaType = FormulaType;
            cf.AlwaysCalculateArray = AlwaysCalculateArray;
            cf.Reference = Reference;
            cf.DataTable2D = DataTable2D;
            cf.DataTableRow = DataTableRow;
            cf.Input1Deleted = Input1Deleted;
            cf.Input2Deleted = Input2Deleted;
            cf.R1 = R1;
            cf.R2 = R2;
            cf.CalculateCell = CalculateCell;
            cf.SharedIndex = SharedIndex;
            cf.Bx = Bx;

            return cf;
        }
    }
}