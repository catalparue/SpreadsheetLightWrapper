using DocumentFormat.OpenXml.Spreadsheet;
using Ups.Toolkit.SpreadsheetLight.Core.misc;

namespace Ups.Toolkit.SpreadsheetLight.Core.workbook
{
    internal class SLCalculationCell
    {
        internal SLCalculationCell()
        {
            SetAllNull();
        }

        internal SLCalculationCell(string CellReference)
        {
            SetAllNull();

            var iRowIndex = -1;
            var iColumnIndex = -1;
            if (SLTool.FormatCellReferenceToRowColumnIndex(CellReference, out iRowIndex, out iColumnIndex))
            {
                RowIndex = iRowIndex;
                ColumnIndex = iColumnIndex;
            }
        }

        internal int RowIndex { get; set; }
        internal int ColumnIndex { get; set; }
        internal int SheetId { get; set; }
        internal bool? InChildChain { get; set; }
        internal bool? NewLevel { get; set; }
        internal bool? NewThread { get; set; }
        internal bool? Array { get; set; }

        private void SetAllNull()
        {
            RowIndex = 1;
            ColumnIndex = 1;
            SheetId = 0;
            InChildChain = null;
            NewLevel = null;
            NewThread = null;
            Array = null;
        }

        internal void FromCalculationCell(CalculationCell cc)
        {
            SetAllNull();

            var iRowIndex = -1;
            var iColumnIndex = -1;
            if (SLTool.FormatCellReferenceToRowColumnIndex(cc.CellReference.Value, out iRowIndex, out iColumnIndex))
            {
                RowIndex = iRowIndex;
                ColumnIndex = iColumnIndex;
            }


            SheetId = cc.SheetId ?? 0;
            if (cc.InChildChain != null) InChildChain = cc.InChildChain.Value;
            if (cc.NewLevel != null) NewLevel = cc.NewLevel.Value;
            if (cc.NewThread != null) NewThread = cc.NewThread.Value;
            if (cc.Array != null) Array = cc.Array.Value;
        }

        internal CalculationCell ToCalculationCell()
        {
            var cc = new CalculationCell();
            cc.CellReference = SLTool.ToCellReference(RowIndex, ColumnIndex);
            cc.SheetId = SheetId;
            if ((InChildChain != null) && InChildChain.Value) cc.InChildChain = InChildChain.Value;
            if ((NewLevel != null) && NewLevel.Value) cc.NewLevel = NewLevel.Value;
            if ((NewThread != null) && NewThread.Value) cc.NewThread = NewThread.Value;
            if ((Array != null) && Array.Value) cc.Array = Array.Value;

            return cc;
        }

        internal SLCalculationCell Clone()
        {
            var cc = new SLCalculationCell();
            cc.RowIndex = RowIndex;
            cc.ColumnIndex = ColumnIndex;
            cc.SheetId = SheetId;
            cc.InChildChain = InChildChain;
            cc.NewLevel = NewLevel;
            cc.NewThread = NewThread;
            cc.Array = Array;

            return cc;
        }
    }
}