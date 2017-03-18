namespace Ups.Toolkit.SpreadsheetLight.Core.misc
{
    /// <summary>
    ///     Statistical information about a worksheet.
    /// </summary>
    public class SLWorksheetStatistics
    {
        internal int iEndColumnIndex;

        internal int iEndRowIndex;

        internal int iNumberOfCells;

        internal int iNumberOfColumns;

        internal int iNumberOfEmptyCells;

        internal int iNumberOfRows;

        internal int iStartColumnIndex;
        internal int iStartRowIndex;

        /// <summary>
        ///     Initializes an instance of SLWorksheetStatistics. But it's quite useless on its own. Use GetWorksheetStatistics()
        ///     of the SLDocument class.
        /// </summary>
        public SLWorksheetStatistics()
        {
            iStartRowIndex = -1;
            iStartColumnIndex = -1;
            iEndRowIndex = -1;
            iEndColumnIndex = -1;
            iNumberOfCells = 0;
            iNumberOfEmptyCells = 0;
            iNumberOfRows = 0;
            iNumberOfColumns = 0;
        }

        /// <summary>
        ///     Index of the first row used. This includes empty rows but might be styled. This returns -1 if the worksheet is
        ///     empty (but check for negative values instead of -1 just in case). This is read-only.
        /// </summary>
        public int StartRowIndex
        {
            get { return iStartRowIndex; }
        }

        /// <summary>
        ///     Index of the first column used. This includes empty columns but might be styled. This returns -1 if the worksheet
        ///     is empty (but check for negative values instead of -1 just in case). This is read-only.
        /// </summary>
        public int StartColumnIndex
        {
            get { return iStartColumnIndex; }
        }

        /// <summary>
        ///     Index of the last row used. This includes empty rows but might be styled. This returns -1 if the worksheet is empty
        ///     (but check for negative values instead of -1 just in case). This is read-only.
        /// </summary>
        public int EndRowIndex
        {
            get { return iEndRowIndex; }
        }

        /// <summary>
        ///     Index of the last column used. This includes empty columns but might be styled. This returns -1 if the worksheet is
        ///     empty (but check for negative values instead of -1 just in case). This is read-only.
        /// </summary>
        public int EndColumnIndex
        {
            get { return iEndColumnIndex; }
        }

        /// <summary>
        ///     Number of cells set in the worksheet. This is read-only.
        /// </summary>
        public int NumberOfCells
        {
            get { return iNumberOfCells; }
        }

        /// <summary>
        ///     Number of cells set in the worksheet that is empty. This could be that a style was set but no cell value given.
        ///     This is read-only.
        /// </summary>
        public int NumberOfEmptyCells
        {
            get { return iNumberOfEmptyCells; }
        }

        /// <summary>
        ///     Number of rows in the worksheet. This includes empty rows (no cells in that row but a row style was applied, or
        ///     that row only has empty cells). This is read-only.
        /// </summary>
        public int NumberOfRows
        {
            get { return iNumberOfRows; }
        }

        /// <summary>
        ///     Number of columns in the worksheet. This includes empty columns (no cells in that column but a column style was
        ///     applied, or that column only has empty cells). This is read-only.
        /// </summary>
        public int NumberOfColumns
        {
            get { return iNumberOfColumns; }
        }
    }
}