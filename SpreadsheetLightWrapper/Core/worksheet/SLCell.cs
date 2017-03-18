using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLightWrapper.Core.misc;

namespace SpreadsheetLightWrapper.Core.worksheet
{
    /// <summary>
    ///     This is for information purposes only! This simulates the DocumentFormat.OpenXml.Spreadsheet.Cell class.
    /// </summary>
    public class SLCell
    {
        /// <summary>
        ///     Access this at your own peril! Only when CellText and NumericValue have to be set together! Probably! You've been
        ///     warned!
        /// </summary>
        internal double fNumericValue;

        private string sCellText;

        internal SLCell()
        {
            SetAllNull();
        }

        /// <summary>
        ///     Indicates if the cell is truly empty. This is read-only.
        /// </summary>
        public bool IsEmpty
        {
            get
            {
                return (CellFormula == null) && (sCellText != null) && (sCellText.Length == 0) && (StyleIndex == 0) &&
                       (DataType == CellValues.Number) && (CellMetaIndex == 0) && (ValueMetaIndex == 0) && !ShowPhonetic;
            }
        }

        //internal CellFormula Formula { get; set; }

        /// <summary>
        ///     Cell formula.
        /// </summary>
        public SLCellFormula CellFormula { get; set; }

        internal bool ToPreserveSpace { get; private set; }

        /// <summary>
        ///     If this is null, the actual value is stored in NumericValue.
        /// </summary>
        public string CellText
        {
            get { return sCellText; }
            set
            {
                sCellText = value;
                ToPreserveSpace = SLTool.ToPreserveSpace(sCellText);

                if (value != null) fNumericValue = 0;
            }
        }

        /// <summary>
        ///     Use this value only when CellText is null.
        /// </summary>
        public double NumericValue
        {
            get { return fNumericValue; }
            set
            {
                fNumericValue = value;

                sCellText = null;
                ToPreserveSpace = false;
            }
        }

        // The logic will be to store boolean, numbers and shared string indices in
        // NumericValue. We'll actually use NumericValue if CellText is null.
        // This will keep memory low since Text will not always be used.
        // Most spreadsheets have numeric data. Consider "1234.56789". That's a 10
        // character string, but is always 8 bytes (?) if stored as a double.
        // In fact, any double is always stored as 8 bytes (thus the memory savings).
        // Plus it seems faster to assign and store a number to a double type than
        // storing the number in string form.

        // So. If CellText is null, it's a number type.
        // If CellText has some string in it, then we use that. And depending on the data type,
        // we'll interpret CellText differently.

        /// <summary>
        ///     Style index.
        /// </summary>
        public uint StyleIndex { get; set; }

        /// <summary>
        ///     Cell data type.
        /// </summary>
        public CellValues DataType { get; set; }

        /// <summary>
        ///     Cell meta index.
        /// </summary>
        public uint CellMetaIndex { get; set; }

        /// <summary>
        ///     Cell value meta index.
        /// </summary>
        public uint ValueMetaIndex { get; set; }

        /// <summary>
        ///     Indicates if phonetic information should be shown.
        /// </summary>
        public bool ShowPhonetic { get; set; }

        internal void SetAllNull()
        {
            //this.Formula = null;
            CellFormula = null;

            ToPreserveSpace = false;
            sCellText = string.Empty;

            fNumericValue = 0;

            StyleIndex = 0;
            DataType = CellValues.Number;
            CellMetaIndex = 0;
            ValueMetaIndex = 0;
            ShowPhonetic = false;
        }

        internal void FromCell(Cell c)
        {
            SetAllNull();

            //if (c.CellFormula != null) this.Formula = (CellFormula)c.CellFormula.CloneNode(true);
            if (c.CellFormula != null)
            {
                CellFormula = new SLCellFormula();
                CellFormula.FromCellFormula(c.CellFormula);
            }

            if (c.StyleIndex != null) StyleIndex = c.StyleIndex.Value;

            if (c.DataType != null) DataType = c.DataType.Value;
            else DataType = CellValues.Number;

            if (c.CellValue != null) CellText = c.CellValue.Text ?? string.Empty;

            double fValue = 0;
            var iValue = 0;
            var bValue = false;
            switch (DataType)
            {
                case CellValues.Number:
                    if (double.TryParse(CellText, NumberStyles.Any, CultureInfo.InvariantCulture, out fValue))
                        NumericValue = fValue;
                    break;
                case CellValues.SharedString:
                    if (int.TryParse(CellText, NumberStyles.Any, CultureInfo.InvariantCulture, out iValue))
                        NumericValue = iValue;
                    break;
                case CellValues.Boolean:
                    if (double.TryParse(CellText, NumberStyles.Any, CultureInfo.InvariantCulture, out fValue))
                        if (fValue > 0.5) NumericValue = 1;
                        else NumericValue = 0;
                    else if (bool.TryParse(CellText, out bValue))
                        if (bValue) NumericValue = 1;
                        else NumericValue = 0;
                    break;
            }

            if (c.CellMetaIndex != null) CellMetaIndex = c.CellMetaIndex.Value;
            if (c.ValueMetaIndex != null) ValueMetaIndex = c.ValueMetaIndex.Value;
            if (c.ShowPhonetic != null) ShowPhonetic = c.ShowPhonetic.Value;
        }

        internal Cell ToCell(string CellReference)
        {
            var c = new Cell();
            //if (this.Formula != null) c.CellFormula = this.Formula;
            if (CellFormula != null) c.CellFormula = CellFormula.ToCellFormula();

            if (CellText != null)
            {
                if (CellText.Length > 0)
                    if (ToPreserveSpace)
                        c.CellValue = new CellValue(CellText)
                        {
                            Space = SpaceProcessingModeValues.Preserve
                        };
                    else
                        c.CellValue = new CellValue(CellText);
            }
            else
            {
                // zero Text length
                if (DataType == CellValues.Number)
                    c.CellValue = new CellValue(NumericValue.ToString(CultureInfo.InvariantCulture));
                else if (DataType == CellValues.SharedString)
                    c.CellValue = new CellValue(NumericValue.ToString("f0", CultureInfo.InvariantCulture));
                else if (DataType == CellValues.Boolean)
                    if (NumericValue > 0.5) c.CellValue = new CellValue("1");
                    else c.CellValue = new CellValue("0");
            }

            c.CellReference = CellReference;
            if (StyleIndex > 0) c.StyleIndex = StyleIndex;
            if (DataType != CellValues.Number) c.DataType = DataType;
            if (CellMetaIndex > 0) c.CellMetaIndex = CellMetaIndex;
            if (ValueMetaIndex > 0) c.ValueMetaIndex = ValueMetaIndex;
            if (ShowPhonetic) c.ShowPhonetic = true;

            return c;
        }

        internal SLCell Clone()
        {
            var cell = new SLCell();
            //if (this.Formula != null) cell.Formula = (CellFormula)this.Formula.CloneNode(true);
            if (CellFormula != null) cell.CellFormula = CellFormula.Clone();
            cell.ToPreserveSpace = ToPreserveSpace;
            cell.sCellText = sCellText;

            cell.fNumericValue = fNumericValue;

            cell.StyleIndex = StyleIndex;
            cell.DataType = DataType;
            cell.CellMetaIndex = CellMetaIndex;
            cell.ValueMetaIndex = ValueMetaIndex;
            cell.ShowPhonetic = ShowPhonetic;

            return cell;
        }
    }
}