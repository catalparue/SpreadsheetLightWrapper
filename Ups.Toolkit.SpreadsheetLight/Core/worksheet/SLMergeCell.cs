using DocumentFormat.OpenXml.Spreadsheet;
using Ups.Toolkit.SpreadsheetLight.Core.misc;

namespace Ups.Toolkit.SpreadsheetLight.Core.worksheet
{
    /// <summary>
    ///     Encapsulates properties and methods for representing a merged cell range. This simulates the
    ///     DocumentFormat.OpenXml.Spreadsheet.MergeCell class.
    ///     The actual merging of cells is done by a SLDocument function. This class is for supporting purposes.
    /// </summary>
    public class SLMergeCell
    {
        internal int iEndColumnIndex;

        internal int iEndRowIndex;

        internal int iStartColumnIndex;
        internal int iStartRowIndex;

        /// <summary>
        ///     Initializes an instance of SLMergeCell.
        /// </summary>
        public SLMergeCell()
        {
            iStartRowIndex = 1;
            iStartColumnIndex = 1;
            iEndRowIndex = 1;
            iEndColumnIndex = 1;
            IsValid = false;
        }

        /// <summary>
        ///     The row index of the top row in the merged cell range. This is read-only.
        /// </summary>
        public int StartRowIndex
        {
            get { return iStartRowIndex; }
        }

        /// <summary>
        ///     The column index of the left column in the merged cell range. This is read-only.
        /// </summary>
        public int StartColumnIndex
        {
            get { return iStartColumnIndex; }
        }

        /// <summary>
        ///     The row index of the bottom row in the merged cell range. This is read-only.
        /// </summary>
        public int EndRowIndex
        {
            get { return iEndRowIndex; }
        }

        /// <summary>
        ///     The column index of the right column in the merged cell range. This is read-only.
        /// </summary>
        public int EndColumnIndex
        {
            get { return iEndColumnIndex; }
        }

        /// <summary>
        ///     Indicates if the merged cell range is valid. This is read-only.
        /// </summary>
        public bool IsValid { get; private set; }

        /// <summary>
        ///     Form a SLMergeCell given a corner cell of the to-be-merged rectangle of cells, and the opposite corner cell. For
        ///     example, the top-left corner cell and the bottom-right corner cell. Or the bottom-left corner cell and the
        ///     top-right corner cell.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the corner cell.</param>
        /// <param name="StartColumnIndex">The column index of the corner cell.</param>
        /// <param name="EndRowIndex">The row index of the opposite corner cell.</param>
        /// <param name="EndColumnIndex">The column index of the opposite corner cell.</param>
        public void FromIndices(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            if (StartRowIndex < EndRowIndex)
            {
                iStartRowIndex = StartRowIndex;
                iEndRowIndex = EndRowIndex;
            }
            else
            {
                iStartRowIndex = EndRowIndex;
                iEndRowIndex = StartRowIndex;
            }

            if (StartColumnIndex < EndColumnIndex)
            {
                iStartColumnIndex = StartColumnIndex;
                iEndColumnIndex = EndColumnIndex;
            }
            else
            {
                iStartColumnIndex = EndColumnIndex;
                iEndColumnIndex = StartColumnIndex;
            }

            if ((iStartRowIndex == iEndRowIndex) && (iStartColumnIndex == iEndColumnIndex))
                IsValid = false;
            else
                IsValid = SLTool.CheckRowColumnIndexLimit(iStartRowIndex, iStartColumnIndex) &&
                          SLTool.CheckRowColumnIndexLimit(iEndRowIndex, iEndColumnIndex);
        }

        /// <summary>
        ///     Form a SLMergeCell from a DocumentFormat.OpenXml.Spreadsheet.MergeCell class.
        /// </summary>
        /// <param name="mc">The source DocumentFormat.OpenXml.Spreadsheet.MergeCell class.</param>
        public void FromMergeCell(MergeCell mc)
        {
            string sStartCell = string.Empty, sEndCell = string.Empty;
            var index = 0;
            bool bStartSuccess = false, bEndSuccess = false;
            IsValid = false;

            if (mc.Reference != null)
            {
                index = mc.Reference.Value.IndexOf(":");
                // if "A1:C3", then the index must be at least at the 3rd position (or index 2)
                if (index >= 2)
                {
                    sStartCell = mc.Reference.Value.Substring(0, index);
                    sEndCell = mc.Reference.Value.Substring(index + 1);

                    bStartSuccess = SLTool.FormatCellReferenceToRowColumnIndex(sStartCell, out iStartRowIndex,
                        out iStartColumnIndex);
                    bEndSuccess = SLTool.FormatCellReferenceToRowColumnIndex(sEndCell, out iEndRowIndex,
                        out iEndColumnIndex);

                    if (bStartSuccess && bEndSuccess)
                        IsValid = true;
                }
            }
        }

        /// <summary>
        ///     Form a DocumentFormat.OpenXml.Spreadsheet.MergeCell class from this SLMergeCell class.
        /// </summary>
        /// <returns>A DocumentFormat.OpenXml.Spreadsheet.MergeCell class.</returns>
        public MergeCell ToMergeCell()
        {
            var mc = new MergeCell();
            var sStartCell = SLTool.ToCellReference(iStartRowIndex, iStartColumnIndex);
            var sEndCell = SLTool.ToCellReference(iEndRowIndex, iEndColumnIndex);
            mc.Reference = string.Format("{0}:{1}", sStartCell, sEndCell);

            return mc;
        }

        internal SLMergeCell Clone()
        {
            var mc = new SLMergeCell();
            mc.iStartRowIndex = iStartRowIndex;
            mc.iStartColumnIndex = iStartColumnIndex;
            mc.iEndRowIndex = iEndRowIndex;
            mc.iEndColumnIndex = iEndColumnIndex;
            mc.IsValid = IsValid;

            return mc;
        }
    }
}