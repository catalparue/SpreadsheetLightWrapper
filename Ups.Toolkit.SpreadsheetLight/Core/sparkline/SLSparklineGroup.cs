using System.Collections.Generic;
using System.Drawing;
using DocumentFormat.OpenXml.Office.Excel;
using Ups.Toolkit.SpreadsheetLight.Core.misc;
using Ups.Toolkit.SpreadsheetLight.Core.style;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace Ups.Toolkit.SpreadsheetLight.Core.sparkline
{
    /// <summary>
    ///     Built-in sparkline styles.
    /// </summary>
    public enum SLSparklineStyle
    {
        /// <summary>
        ///     Accent 1 Darker 50%
        /// </summary>
        Accent1Darker50Percent = 0,

        /// <summary>
        ///     Accent 2 Darker 50%
        /// </summary>
        Accent2Darker50Percent,

        /// <summary>
        ///     Accent 3 Darker 50%
        /// </summary>
        Accent3Darker50Percent,

        /// <summary>
        ///     Accent 4 Darker 50%
        /// </summary>
        Accent4Darker50Percent,

        /// <summary>
        ///     Accent 5 Darker 50%
        /// </summary>
        Accent5Darker50Percent,

        /// <summary>
        ///     Accent 6 Darker 50%
        /// </summary>
        Accent6Darker50Percent,

        /// <summary>
        ///     Accent 1 Darker 25%
        /// </summary>
        Accent1Darker25Percent,

        /// <summary>
        ///     Accent 2 Darker 25%
        /// </summary>
        Accent2Darker25Percent,

        /// <summary>
        ///     Accent 3 Darker 25%
        /// </summary>
        Accent3Darker25Percent,

        /// <summary>
        ///     Accent 4 Darker 25%
        /// </summary>
        Accent4Darker25Percent,

        /// <summary>
        ///     Accent 5 Darker 25%
        /// </summary>
        Accent5Darker25Percent,

        /// <summary>
        ///     Accent 6 Darker 25%
        /// </summary>
        Accent6Darker25Percent,

        /// <summary>
        ///     Accent 1
        /// </summary>
        Accent1,

        /// <summary>
        ///     Accent 2
        /// </summary>
        Accent2,

        /// <summary>
        ///     Accent 3
        /// </summary>
        Accent3,

        /// <summary>
        ///     Accent 4
        /// </summary>
        Accent4,

        /// <summary>
        ///     Accent 5
        /// </summary>
        Accent5,

        /// <summary>
        ///     Accent 6
        /// </summary>
        Accent6,

        /// <summary>
        ///     Accent 1 Lighter 40%
        /// </summary>
        Accent1Lighter40Percent,

        /// <summary>
        ///     Accent 2 Lighter 40%
        /// </summary>
        Accent2Lighter40Percent,

        /// <summary>
        ///     Accent 3 Lighter 40%
        /// </summary>
        Accent3Lighter40Percent,

        /// <summary>
        ///     Accent 4 Lighter 40%
        /// </summary>
        Accent4Lighter40Percent,

        /// <summary>
        ///     Accent 5 Lighter 40%
        /// </summary>
        Accent5Lighter40Percent,

        /// <summary>
        ///     Accent 6 Lighter 40%
        /// </summary>
        Accent6Lighter40Percent,

        /// <summary>
        ///     Dark #1
        /// </summary>
        Dark1,

        /// <summary>
        ///     Dark #2
        /// </summary>
        Dark2,

        /// <summary>
        ///     Dark #3
        /// </summary>
        Dark3,

        /// <summary>
        ///     Dark #4
        /// </summary>
        Dark4,

        /// <summary>
        ///     Dark #5
        /// </summary>
        Dark5,

        /// <summary>
        ///     Dark #6
        /// </summary>
        Dark6,

        /// <summary>
        ///     Colorful #1
        /// </summary>
        Colorful1,

        /// <summary>
        ///     Colorful #2
        /// </summary>
        Colorful2,

        /// <summary>
        ///     Colorful #3
        /// </summary>
        Colorful3,

        /// <summary>
        ///     Colorful #4
        /// </summary>
        Colorful4,

        /// <summary>
        ///     Colorful #5
        /// </summary>
        Colorful5,

        /// <summary>
        ///     Colorful #6
        /// </summary>
        Colorful6
    }

    /// <summary>
    ///     Encapsulates properties and methods for specifying sparklines.
    ///     This simulates the DocumentFormat.OpenXml.Office2010.Excel.SparklineGroup class.
    /// </summary>
    public class SLSparklineGroup
    {
        internal int DateEndColumnIndex;
        internal int DateEndRowIndex;
        internal int DateStartColumnIndex;
        internal int DateStartRowIndex;

        internal string DateWorksheetName;

        internal decimal decLineWeight;
        internal int EndColumnIndex;
        internal int EndRowIndex;
        internal List<Color> listIndexedColors;
        internal List<Color> listThemeColors;
        internal int StartColumnIndex;
        internal int StartRowIndex;

        // these are only used for setting location. They're not synchronised if the individual
        // sparkline changes cell references.
        internal string WorksheetName;

        internal SLSparklineGroup(List<Color> ThemeColors, List<Color> IndexedColors)
        {
            int i;
            listThemeColors = new List<Color>();
            for (i = 0; i < ThemeColors.Count; ++i)
                listThemeColors.Add(ThemeColors[i]);

            listIndexedColors = new List<Color>();
            for (i = 0; i < IndexedColors.Count; ++i)
                listIndexedColors.Add(IndexedColors[i]);

            SetAllNull();
        }

        /// <summary>
        ///     The color for the main sparkline series.
        /// </summary>
        public SLColor SeriesColor { get; set; }

        /// <summary>
        ///     The color for negative points.
        /// </summary>
        public SLColor NegativeColor { get; set; }

        /// <summary>
        ///     The color for the axis.
        /// </summary>
        public SLColor AxisColor { get; set; }

        /// <summary>
        ///     The color for markers.
        /// </summary>
        public SLColor MarkersColor { get; set; }

        /// <summary>
        ///     The color for the first point.
        /// </summary>
        public SLColor FirstMarkerColor { get; set; }

        /// <summary>
        ///     The color for the last point.
        /// </summary>
        public SLColor LastMarkerColor { get; set; }

        /// <summary>
        ///     The color for the high point.
        /// </summary>
        public SLColor HighMarkerColor { get; set; }

        /// <summary>
        ///     The color for the low point.
        /// </summary>
        public SLColor LowMarkerColor { get; set; }

        internal double ManualMax { get; set; }
        internal X14.SparklineAxisMinMaxValues MaxAxisType { get; set; }
        internal double ManualMin { get; set; }
        internal X14.SparklineAxisMinMaxValues MinAxisType { get; set; }

        /// <summary>
        ///     Line weight for the sparkline group in points, ranging from 0 pt to 1584 pt (both inclusive).
        /// </summary>
        public decimal LineWeight
        {
            get { return decLineWeight; }
            set
            {
                decLineWeight = value;
                if (decLineWeight < 0) decLineWeight = 0;
                if (decLineWeight > 1584) decLineWeight = 1584;
            }
        }

        /// <summary>
        ///     The type of sparkline. Use "Stacked" for "Win/Loss".
        /// </summary>
        public X14.SparklineTypeValues Type { get; set; }

        internal bool DateAxis { get; set; }

        /// <summary>
        ///     The default is to show empty cells with a gap.
        /// </summary>
        public X14.DisplayBlanksAsValues ShowEmptyCellsAs { get; set; }

        /// <summary>
        ///     Specifies if markers are shown.
        /// </summary>
        public bool ShowMarkers { get; set; }

        /// <summary>
        ///     Specifies if the high point is shown.
        /// </summary>
        public bool ShowHighPoint { get; set; }

        /// <summary>
        ///     Specifies if the low point is shown.
        /// </summary>
        public bool ShowLowPoint { get; set; }

        /// <summary>
        ///     Specifies if the first point is shown.
        /// </summary>
        public bool ShowFirstPoint { get; set; }

        /// <summary>
        ///     Specifies if the last point is shown.
        /// </summary>
        public bool ShowLastPoint { get; set; }

        /// <summary>
        ///     Specifies if negative points are shown.
        /// </summary>
        public bool ShowNegativePoints { get; set; }

        /// <summary>
        ///     Specifies is the horizontal axis is shown. This only appears if there's sparkline data crossing the zero point.
        /// </summary>
        public bool ShowAxis { get; set; }

        /// <summary>
        ///     Specifies if hidden data is shown.
        /// </summary>
        public bool ShowHiddenData { get; set; }

        /// <summary>
        ///     Plot data right-to-left.
        /// </summary>
        public bool RightToLeft { get; set; }

        // supposed to contain less than 2^31 sparklines. But I'm not gonna enforce this...
        // See documentation on CT_Sparklines for this.
        internal List<SLSparkline> Sparklines { get; set; }

        private void SetAllNull()
        {
            WorksheetName = string.Empty;
            StartRowIndex = 1;
            StartColumnIndex = 1;
            EndRowIndex = 1;
            EndColumnIndex = 1;

            SeriesColor = new SLColor(listThemeColors, listIndexedColors);
            NegativeColor = new SLColor(listThemeColors, listIndexedColors);
            AxisColor = new SLColor(listThemeColors, listIndexedColors);
            MarkersColor = new SLColor(listThemeColors, listIndexedColors);
            FirstMarkerColor = new SLColor(listThemeColors, listIndexedColors);
            LastMarkerColor = new SLColor(listThemeColors, listIndexedColors);
            HighMarkerColor = new SLColor(listThemeColors, listIndexedColors);
            LowMarkerColor = new SLColor(listThemeColors, listIndexedColors);

            ManualMax = 0;
            MaxAxisType = X14.SparklineAxisMinMaxValues.Individual;
            ManualMin = 0;
            MinAxisType = X14.SparklineAxisMinMaxValues.Individual;

            decLineWeight = 0.75m;

            Type = X14.SparklineTypeValues.Line;

            DateWorksheetName = string.Empty;
            DateStartRowIndex = 1;
            DateStartColumnIndex = 1;
            DateEndRowIndex = 1;
            DateEndColumnIndex = 1;
            DateAxis = false;

            ShowEmptyCellsAs = X14.DisplayBlanksAsValues.Gap;

            ShowMarkers = false;
            ShowHighPoint = false;
            ShowLowPoint = false;
            ShowFirstPoint = false;
            ShowLastPoint = false;
            ShowNegativePoints = false;
            ShowAxis = false;
            ShowHiddenData = false;
            RightToLeft = false;

            Sparklines = new List<SLSparkline>();
        }

        /// <summary>
        ///     Set the location of the sparkline group given a cell reference. Use this if your data source is either 1 row of
        ///     cells or 1 column of cells.
        /// </summary>
        /// <param name="CellReference">The cell reference such as "A1".</param>
        public void SetLocation(string CellReference)
        {
            // in case developers copy straight from the Excel dialog box...
            var sCellReference = CellReference.Replace("$", "");

            var iRowIndex = -1;
            var iColumnIndex = -1;
            if (!SLTool.FormatCellReferenceToRowColumnIndex(sCellReference, out iRowIndex, out iColumnIndex))
            {
                iRowIndex = -1;
                iColumnIndex = -1;
            }

            SetLocation(iRowIndex, iColumnIndex, iRowIndex, iColumnIndex, true);
        }

        /// <summary>
        ///     Set the location of the sparkline group given cell references of opposite cells in a cell range.
        ///     Note that the cell range has to be a 1-dimensional vector, meaning it's either a single row or single column.
        ///     Note also that the length of the vector must be equal to either the number of rows or number of columns in the data
        ///     source range.
        /// </summary>
        /// <param name="StartCellReference">
        ///     The cell reference of the start cell of the location cell range, such as "A1". This is
        ///     either the top-most or left-most cell.
        /// </param>
        /// <param name="EndCellReference">
        ///     The cell reference of the end cell of the location cell range, such as "A1". This is
        ///     either the bottom-most or right-most cell.
        /// </param>
        public void SetLocation(string StartCellReference, string EndCellReference)
        {
            // in case developers copy straight from the Excel dialog box...
            var sStartCellReference = StartCellReference.Replace("$", "");
            var sEndCellReference = EndCellReference.Replace("$", "");

            var iStartRowIndex = -1;
            var iStartColumnIndex = -1;
            var iEndRowIndex = -1;
            var iEndColumnIndex = -1;
            if (
                !SLTool.FormatCellReferenceToRowColumnIndex(sStartCellReference, out iStartRowIndex,
                    out iStartColumnIndex))
            {
                iStartRowIndex = -1;
                iStartColumnIndex = -1;
            }
            if (!SLTool.FormatCellReferenceToRowColumnIndex(sEndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                iEndRowIndex = -1;
                iEndColumnIndex = -1;
            }

            SetLocation(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, true);
        }

        /// <summary>
        ///     Set the location of the sparkline group given cell references of opposite cells in a cell range.
        ///     Note that the cell range has to be a 1-dimensional vector, meaning it's either a single row or single column.
        ///     Note also that the length of the vector must be equal to either the number of rows or number of columns in the data
        ///     source range.
        /// </summary>
        /// <param name="StartCellReference">
        ///     The cell reference of the start cell of the location cell range, such as "A1". This is
        ///     either the top-most or left-most cell.
        /// </param>
        /// <param name="EndCellReference">
        ///     The cell reference of the end cell of the location cell range, such as "A1". This is
        ///     either the bottom-most or right-most cell.
        /// </param>
        /// <param name="RowsAsDataSeries">
        ///     True if the data source has its series in rows. False if it's in columns. This only
        ///     comes into play if the data source has the same number of rows as its columns.
        /// </param>
        public void SetLocation(string StartCellReference, string EndCellReference, bool RowsAsDataSeries)
        {
            // in case developers copy straight from the Excel dialog box...
            var sStartCellReference = StartCellReference.Replace("$", "");
            var sEndCellReference = EndCellReference.Replace("$", "");

            var iStartRowIndex = -1;
            var iStartColumnIndex = -1;
            var iEndRowIndex = -1;
            var iEndColumnIndex = -1;
            if (
                !SLTool.FormatCellReferenceToRowColumnIndex(sStartCellReference, out iStartRowIndex,
                    out iStartColumnIndex))
            {
                iStartRowIndex = -1;
                iStartColumnIndex = -1;
            }
            if (!SLTool.FormatCellReferenceToRowColumnIndex(sEndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                iEndRowIndex = -1;
                iEndColumnIndex = -1;
            }

            SetLocation(iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex, RowsAsDataSeries);
        }

        /// <summary>
        ///     Set the location of the sparkline group given a row and column index. Use this if your data source is either 1 row
        ///     of cells or 1 column of cells.
        /// </summary>
        /// <param name="RowIndex">The row index.</param>
        /// <param name="ColumnIndex">The column index.</param>
        public void SetLocation(int RowIndex, int ColumnIndex)
        {
            SetLocation(RowIndex, ColumnIndex, RowIndex, ColumnIndex, true);
        }

        /// <summary>
        ///     Set the location of the sparkline group given row and column indices of opposite cells in a cell range.
        ///     Note that the cell range has to be a 1-dimensional vector, meaning it's either a single row or single column.
        ///     Note also that the length of the vector must be equal to either the number of rows or number of columns in the data
        ///     source range.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start row. This is typically the top row.</param>
        /// <param name="StartColumnIndex">The column index of the start column. This is typically the left-most column.</param>
        /// <param name="EndRowIndex">The row index of the end row. This is typically the bottom row.</param>
        /// <param name="EndColumnIndex">The column index of the end column. This is typically the right-most column.</param>
        public void SetLocation(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            SetLocation(StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex, true);
        }

        /// <summary>
        ///     Set the location of the sparkline group given row and column indices of opposite cells in a cell range.
        ///     Note that the cell range has to be a 1-dimensional vector, meaning it's either a single row or single column.
        ///     Note also that the length of the vector must be equal to either the number of rows or number of columns in the data
        ///     source range.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start row. This is typically the top row.</param>
        /// <param name="StartColumnIndex">The column index of the start column. This is typically the left-most column.</param>
        /// <param name="EndRowIndex">The row index of the end row. This is typically the bottom row.</param>
        /// <param name="EndColumnIndex">The column index of the end column. This is typically the right-most column.</param>
        /// <param name="RowsAsDataSeries">
        ///     True if the data source has its series in rows. False if it's in columns. This only
        ///     comes into play if the data source has the same number of rows as its columns.
        /// </param>
        public void SetLocation(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex,
            bool RowsAsDataSeries)
        {
            int iStartRowIndex = 1, iEndRowIndex = 1, iStartColumnIndex = 1, iEndColumnIndex = 1;
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

            if (iStartRowIndex < 1) iStartRowIndex = 1;
            if (iStartColumnIndex < 1) iStartColumnIndex = 1;
            if (iEndRowIndex > SLConstants.RowLimit) iEndRowIndex = SLConstants.RowLimit;
            if (iEndColumnIndex > SLConstants.ColumnLimit) iEndColumnIndex = SLConstants.ColumnLimit;

            var iLocationRowDimension = iEndRowIndex - iStartRowIndex + 1;
            var iLocationColumnDimension = iEndColumnIndex - iStartColumnIndex + 1;

            // either the location row or column dimension must be 1. One of them has to be 1
            // for the location range to be valid. If there's an error, we'll just "shorten"
            // the smaller of the 2 dimensions. The Excel user interface has error dialog boxes
            // to warn the user. We don't have this luxury, so we'll make the best of things...
            if ((iLocationRowDimension != 1) && (iLocationColumnDimension != 1))
                if (iLocationRowDimension < iLocationColumnDimension)
                {
                    iEndRowIndex = iStartRowIndex;
                    iLocationRowDimension = 1;
                }
                else
                {
                    iEndColumnIndex = iStartColumnIndex;
                    iLocationColumnDimension = 1;
                }

            var iDataRowDimension = this.EndRowIndex - this.StartRowIndex + 1;
            var iDataColumnDimension = this.EndColumnIndex - this.StartColumnIndex + 1;

            var bRowsAsDataSeries = true;
            var iMaxLocationDimension = 1;
            if (iLocationRowDimension >= iLocationColumnDimension)
            {
                iMaxLocationDimension = iLocationRowDimension;
                bRowsAsDataSeries = true;
            }
            else
            {
                iMaxLocationDimension = iLocationColumnDimension;
                bRowsAsDataSeries = false;
            }

            // If the data source has the same number of rows as its columns, the "default" is to use rows as data series,
            // unless otherwise stated. This is the "otherwise stated" part.
            if (iDataRowDimension == iDataColumnDimension)
                bRowsAsDataSeries = RowsAsDataSeries;

            // Furthermore, the "length" of the location range has to be either equal to
            // the data source range's row dimension or column dimension.
            // This is how Excel determines whether to use rows or columns as data series.

            int index;
            SLSparkline spk;
            if (iMaxLocationDimension == iDataRowDimension)
                for (index = 0; index < iMaxLocationDimension; ++index)
                {
                    spk = new SLSparkline();
                    spk.WorksheetName = WorksheetName;
                    spk.StartRowIndex = index + this.StartRowIndex;
                    spk.EndRowIndex = spk.StartRowIndex;
                    spk.StartColumnIndex = this.StartColumnIndex;
                    spk.EndColumnIndex = this.EndColumnIndex;

                    if (bRowsAsDataSeries)
                    {
                        spk.LocationRowIndex = index + iStartRowIndex;
                        spk.LocationColumnIndex = iStartColumnIndex;
                    }
                    else
                    {
                        spk.LocationRowIndex = iStartRowIndex;
                        spk.LocationColumnIndex = index + iStartColumnIndex;
                    }

                    Sparklines.Add(spk);
                }
            else if (iMaxLocationDimension == iDataColumnDimension)
                for (index = 0; index < iMaxLocationDimension; ++index)
                {
                    spk = new SLSparkline();
                    spk.WorksheetName = WorksheetName;
                    spk.StartRowIndex = this.StartRowIndex;
                    spk.EndRowIndex = this.EndRowIndex;
                    spk.StartColumnIndex = index + this.StartColumnIndex;
                    spk.EndColumnIndex = spk.StartColumnIndex;

                    if (bRowsAsDataSeries)
                    {
                        spk.LocationRowIndex = index + iStartRowIndex;
                        spk.LocationColumnIndex = iStartColumnIndex;
                    }
                    else
                    {
                        spk.LocationRowIndex = iStartRowIndex;
                        spk.LocationColumnIndex = index + iStartColumnIndex;
                    }

                    Sparklines.Add(spk);
                }
        }

        /// <summary>
        ///     Set the horizontal axis as general axis type.
        /// </summary>
        public void SetGeneralAxis()
        {
            DateAxis = false;
        }

        /// <summary>
        ///     Set the horizontal axis as date axis type, given a cell range containing the date values.
        ///     Note that this means the cell range is a 1-dimensional vector, meaning it's a single row or single column.
        ///     Note also that this probably means the length of the vector is the same as your location cell range.
        /// </summary>
        /// <param name="StartCellReference">
        ///     The cell reference of the start cell of the date cell range, such as "A1". This is
        ///     either the top-most or left-most cell.
        /// </param>
        /// <param name="EndCellReference">
        ///     The cell reference of the end cell of the date cell range, such as "A1". This is either
        ///     the bottom-most or right-most cell.
        /// </param>
        public void SetDateAxis(string StartCellReference, string EndCellReference)
        {
            var iStartRowIndex = -1;
            var iStartColumnIndex = -1;
            var iEndRowIndex = -1;
            var iEndColumnIndex = -1;
            if (
                !SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex,
                    out iStartColumnIndex))
            {
                iStartRowIndex = -1;
                iStartColumnIndex = -1;
            }
            if (!SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                iEndRowIndex = -1;
                iEndColumnIndex = -1;
            }

            SetDateAxis(WorksheetName, iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex);
        }

        /// <summary>
        ///     Set the horizontal axis as date axis type, given a worksheet name and a cell range containing the date values.
        ///     Note that this means the cell range is a 1-dimensional vector, meaning it's a single row or single column.
        ///     Note also that this probably means the length of the vector is the same as your location cell range.
        /// </summary>
        /// <param name="WorksheetName">The name of the source worksheet.</param>
        /// <param name="StartCellReference">
        ///     The cell reference of the start cell of the date cell range, such as "A1". This is
        ///     either the top-most or left-most cell.
        /// </param>
        /// <param name="EndCellReference">
        ///     The cell reference of the end cell of the date cell range, such as "A1". This is either
        ///     the bottom-most or right-most cell.
        /// </param>
        public void SetDateAxis(string WorksheetName, string StartCellReference, string EndCellReference)
        {
            var iStartRowIndex = -1;
            var iStartColumnIndex = -1;
            var iEndRowIndex = -1;
            var iEndColumnIndex = -1;
            if (
                !SLTool.FormatCellReferenceToRowColumnIndex(StartCellReference, out iStartRowIndex,
                    out iStartColumnIndex))
            {
                iStartRowIndex = -1;
                iStartColumnIndex = -1;
            }
            if (!SLTool.FormatCellReferenceToRowColumnIndex(EndCellReference, out iEndRowIndex, out iEndColumnIndex))
            {
                iEndRowIndex = -1;
                iEndColumnIndex = -1;
            }

            SetDateAxis(WorksheetName, iStartRowIndex, iStartColumnIndex, iEndRowIndex, iEndColumnIndex);
        }

        /// <summary>
        ///     Set the horizontal axis as date axis type, given row and column indices of opposite cells in a cell range
        ///     containing the date values.
        ///     Note that this means the cell range is a 1-dimensional vector, meaning it's a single row or single column.
        ///     Note also that this probably means the length of the vector is the same as your location cell range.
        /// </summary>
        /// <param name="StartRowIndex">The row index of the start row. This is typically the top row.</param>
        /// <param name="StartColumnIndex">The column index of the start column. This is typically the left-most column.</param>
        /// <param name="EndRowIndex">The row index of the end row. This is typically the bottom row.</param>
        /// <param name="EndColumnIndex">The column index of the end column. This is typically the right-most column.</param>
        public void SetDateAxis(int StartRowIndex, int StartColumnIndex, int EndRowIndex, int EndColumnIndex)
        {
            SetDateAxis(WorksheetName, StartRowIndex, StartColumnIndex, EndRowIndex, EndColumnIndex);
        }

        /// <summary>
        ///     Set the horizontal axis as date axis type, given a worksheet name, and row and column indices of opposite cells in
        ///     a cell range containing the date values.
        ///     Note that this means the cell range is a 1-dimensional vector, meaning it's a single row or single column.
        ///     Note also that this probably means the length of the vector is the same as your location cell range.
        /// </summary>
        /// <param name="WorksheetName">The name of the source worksheet.</param>
        /// <param name="StartRowIndex">The row index of the start row. This is typically the top row.</param>
        /// <param name="StartColumnIndex">The column index of the start column. This is typically the left-most column.</param>
        /// <param name="EndRowIndex">The row index of the end row. This is typically the bottom row.</param>
        /// <param name="EndColumnIndex">The column index of the end column. This is typically the right-most column.</param>
        public void SetDateAxis(string WorksheetName, int StartRowIndex, int StartColumnIndex, int EndRowIndex,
            int EndColumnIndex)
        {
            DateAxis = true;

            int iStartRowIndex = 1, iEndRowIndex = 1, iStartColumnIndex = 1, iEndColumnIndex = 1;
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

            if (iStartRowIndex < 1) iStartRowIndex = 1;
            if (iStartColumnIndex < 1) iStartColumnIndex = 1;
            if (iEndRowIndex > SLConstants.RowLimit) iEndRowIndex = SLConstants.RowLimit;
            if (iEndColumnIndex > SLConstants.ColumnLimit) iEndColumnIndex = SLConstants.ColumnLimit;

            DateWorksheetName = WorksheetName;
            DateStartRowIndex = iStartRowIndex;
            DateStartColumnIndex = iStartColumnIndex;
            DateEndRowIndex = iEndRowIndex;
            DateEndColumnIndex = iEndColumnIndex;
        }

        /// <summary>
        ///     Set automatic minimum value for the vertical axis for the entire sparkline group.
        /// </summary>
        public void SetAutomaticMinimumValue()
        {
            MinAxisType = X14.SparklineAxisMinMaxValues.Individual;
            ManualMin = 0;
        }

        /// <summary>
        ///     Set the same minimum value for the vertical axis for the entire sparkline group.
        /// </summary>
        public void SetSameMinimumValue()
        {
            MinAxisType = X14.SparklineAxisMinMaxValues.Group;
            ManualMin = 0;
        }

        /// <summary>
        ///     Set a custom minimum value for the vertical axis for the entire sparkline group.
        /// </summary>
        /// <param name="MinValue">The custom minimum value.</param>
        public void SetCustomMinimumValue(double MinValue)
        {
            MinAxisType = X14.SparklineAxisMinMaxValues.Custom;
            ManualMin = MinValue;
        }

        /// <summary>
        ///     Set automatic maximum value for the vertical axis for the entire sparkline group.
        /// </summary>
        public void SetAutomaticMaximumValue()
        {
            MaxAxisType = X14.SparklineAxisMinMaxValues.Individual;
            ManualMax = 0;
        }

        /// <summary>
        ///     Set the same maximum value for the vertical axis for the entire sparkline group.
        /// </summary>
        public void SetSameMaximumValue()
        {
            MaxAxisType = X14.SparklineAxisMinMaxValues.Group;
            ManualMax = 0;
        }

        /// <summary>
        ///     Set a custom maximum value for the vertical axis for the entire sparkline group.
        /// </summary>
        /// <param name="MaxValue">The custom maximum value.</param>
        public void SetCustomMaximumValue(double MaxValue)
        {
            MaxAxisType = X14.SparklineAxisMinMaxValues.Custom;
            ManualMax = MaxValue;
        }

        /// <summary>
        ///     Set the sparkline style.
        /// </summary>
        /// <param name="Style">A built-in sparkline style.</param>
        public void SetSparklineStyle(SLSparklineStyle Style)
        {
            switch (Style)
            {
                case SLSparklineStyle.Accent1Darker50Percent:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.499984740745262);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.499984740745262);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, 0.39997558519241921);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, 0.39997558519241921);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color);
                    break;
                case SLSparklineStyle.Accent2Darker50Percent:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.499984740745262);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.499984740745262);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, 0.39997558519241921);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, 0.39997558519241921);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color);
                    break;
                case SLSparklineStyle.Accent3Darker50Percent:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.499984740745262);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.499984740745262);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, 0.39997558519241921);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, 0.39997558519241921);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color);
                    break;
                case SLSparklineStyle.Accent4Darker50Percent:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.499984740745262);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.499984740745262);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, 0.39997558519241921);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, 0.39997558519241921);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color);
                    break;
                case SLSparklineStyle.Accent5Darker50Percent:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.499984740745262);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.499984740745262);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, 0.39997558519241921);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, 0.39997558519241921);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color);
                    break;
                case SLSparklineStyle.Accent6Darker50Percent:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.499984740745262);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.499984740745262);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, 0.39997558519241921);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, 0.39997558519241921);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color);
                    break;
                case SLSparklineStyle.Accent1Darker25Percent:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.249977111117893);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.249977111117893);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.249977111117893);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.249977111117893);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.249977111117893);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.249977111117893);
                    break;
                case SLSparklineStyle.Accent2Darker25Percent:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.249977111117893);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.249977111117893);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.249977111117893);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.249977111117893);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.249977111117893);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.249977111117893);
                    break;
                case SLSparklineStyle.Accent3Darker25Percent:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.249977111117893);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.249977111117893);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.249977111117893);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.249977111117893);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.249977111117893);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.249977111117893);
                    break;
                case SLSparklineStyle.Accent4Darker25Percent:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.249977111117893);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.249977111117893);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.249977111117893);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.249977111117893);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.249977111117893);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.249977111117893);
                    break;
                case SLSparklineStyle.Accent5Darker25Percent:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.249977111117893);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.249977111117893);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.249977111117893);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.249977111117893);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.249977111117893);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.249977111117893);
                    break;
                case SLSparklineStyle.Accent6Darker25Percent:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.249977111117893);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.249977111117893);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.249977111117893);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.249977111117893);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.249977111117893);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.249977111117893);
                    break;
                case SLSparklineStyle.Accent1:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.249977111117893);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.249977111117893);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.249977111117893);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.249977111117893);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.249977111117893);
                    break;
                case SLSparklineStyle.Accent2:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.249977111117893);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.249977111117893);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.249977111117893);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.249977111117893);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.249977111117893);
                    break;
                case SLSparklineStyle.Accent3:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.249977111117893);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.249977111117893);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.249977111117893);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.249977111117893);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.249977111117893);
                    break;
                case SLSparklineStyle.Accent4:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.249977111117893);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.249977111117893);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.249977111117893);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.249977111117893);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.249977111117893);
                    break;
                case SLSparklineStyle.Accent5:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.249977111117893);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.249977111117893);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.249977111117893);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.249977111117893);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.249977111117893);
                    break;
                case SLSparklineStyle.Accent6:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.249977111117893);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.249977111117893);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.249977111117893);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.249977111117893);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.249977111117893);
                    break;
                case SLSparklineStyle.Accent1Lighter40Percent:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, 0.39997558519241921);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Light1Color, -0.499984740745262);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, 0.79998168889431442);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.249977111117893);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.249977111117893);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.499984740745262);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color, -0.499984740745262);
                    break;
                case SLSparklineStyle.Accent2Lighter40Percent:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, 0.39997558519241921);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Light1Color, -0.499984740745262);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, 0.79998168889431442);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.249977111117893);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.249977111117893);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.499984740745262);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color, -0.499984740745262);
                    break;
                case SLSparklineStyle.Accent3Lighter40Percent:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, 0.39997558519241921);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Light1Color, -0.499984740745262);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, 0.79998168889431442);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.249977111117893);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.249977111117893);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.499984740745262);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color, -0.499984740745262);
                    break;
                case SLSparklineStyle.Accent4Lighter40Percent:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, 0.39997558519241921);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Light1Color, -0.499984740745262);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, 0.79998168889431442);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.249977111117893);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.249977111117893);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.499984740745262);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color, -0.499984740745262);
                    break;
                case SLSparklineStyle.Accent5Lighter40Percent:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, 0.39997558519241921);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Light1Color, -0.499984740745262);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, 0.79998168889431442);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.249977111117893);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.249977111117893);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.499984740745262);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color, -0.499984740745262);
                    break;
                case SLSparklineStyle.Accent6Lighter40Percent:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, 0.39997558519241921);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Light1Color, -0.499984740745262);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, 0.79998168889431442);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.249977111117893);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.249977111117893);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.499984740745262);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color, -0.499984740745262);
                    break;
                case SLSparklineStyle.Dark1:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Dark1Color, 0.499984740745262);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Dark1Color, 0.249977111117893);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Dark1Color, 0.249977111117893);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Dark1Color, 0.249977111117893);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Dark1Color, 0.249977111117893);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Dark1Color, 0.249977111117893);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Dark1Color, 0.249977111117893);
                    break;
                case SLSparklineStyle.Dark2:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Dark1Color, 0.34998626667073579);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Light1Color, -0.249977111117893);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Light1Color, -0.249977111117893);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Light1Color, -0.249977111117893);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Light1Color, -0.249977111117893);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Light1Color, -0.249977111117893);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Light1Color, -0.249977111117893);
                    break;
                case SLSparklineStyle.Dark3:
                    SeriesColor.Color = Color.FromArgb(0xFF, 0x32, 0x32, 0x32);
                    NegativeColor.Color = Color.FromArgb(0xFF, 0xD0, 0, 0);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.Color = Color.FromArgb(0xFF, 0xD0, 0, 0);
                    FirstMarkerColor.Color = Color.FromArgb(0xFF, 0xD0, 0, 0);
                    LastMarkerColor.Color = Color.FromArgb(0xFF, 0xD0, 0, 0);
                    HighMarkerColor.Color = Color.FromArgb(0xFF, 0xD0, 0, 0);
                    LowMarkerColor.Color = Color.FromArgb(0xFF, 0xD0, 0, 0);
                    break;
                case SLSparklineStyle.Dark4:
                    SeriesColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    NegativeColor.Color = Color.FromArgb(0xFF, 0, 0x70, 0xC0);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.Color = Color.FromArgb(0xFF, 0, 0x70, 0xC0);
                    FirstMarkerColor.Color = Color.FromArgb(0xFF, 0, 0x70, 0xC0);
                    LastMarkerColor.Color = Color.FromArgb(0xFF, 0, 0x70, 0xC0);
                    HighMarkerColor.Color = Color.FromArgb(0xFF, 0, 0x70, 0xC0);
                    LowMarkerColor.Color = Color.FromArgb(0xFF, 0, 0x70, 0xC0);
                    break;
                case SLSparklineStyle.Dark5:
                    SeriesColor.Color = Color.FromArgb(0xFF, 0x37, 0x60, 0x92);
                    NegativeColor.Color = Color.FromArgb(0xFF, 0xD0, 0, 0);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.Color = Color.FromArgb(0xFF, 0xD0, 0, 0);
                    FirstMarkerColor.Color = Color.FromArgb(0xFF, 0xD0, 0, 0);
                    LastMarkerColor.Color = Color.FromArgb(0xFF, 0xD0, 0, 0);
                    HighMarkerColor.Color = Color.FromArgb(0xFF, 0xD0, 0, 0);
                    LowMarkerColor.Color = Color.FromArgb(0xFF, 0xD0, 0, 0);
                    break;
                case SLSparklineStyle.Dark6:
                    SeriesColor.Color = Color.FromArgb(0xFF, 0, 0x70, 0xC0);
                    NegativeColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    FirstMarkerColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    LastMarkerColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    HighMarkerColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    LowMarkerColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    break;
                case SLSparklineStyle.Colorful1:
                    SeriesColor.Color = Color.FromArgb(0xFF, 0x5F, 0x5F, 0x5F);
                    NegativeColor.Color = Color.FromArgb(0xFF, 0xFF, 0xB6, 0x20);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.Color = Color.FromArgb(0xFF, 0xD7, 0, 0x77);
                    FirstMarkerColor.Color = Color.FromArgb(0xFF, 0x56, 0x87, 0xC2);
                    LastMarkerColor.Color = Color.FromArgb(0xFF, 0x35, 0x9C, 0xEB);
                    HighMarkerColor.Color = Color.FromArgb(0xFF, 0x56, 0xBE, 0x79);
                    LowMarkerColor.Color = Color.FromArgb(0xFF, 0xFF, 0x50, 0x55);
                    break;
                case SLSparklineStyle.Colorful2:
                    SeriesColor.Color = Color.FromArgb(0xFF, 0x56, 0x87, 0xC2);
                    NegativeColor.Color = Color.FromArgb(0xFF, 0xFF, 0xB6, 0x20);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.Color = Color.FromArgb(0xFF, 0xD7, 0, 0x77);
                    FirstMarkerColor.Color = Color.FromArgb(0xFF, 0x77, 0x77, 0x77);
                    LastMarkerColor.Color = Color.FromArgb(0xFF, 0x35, 0x9C, 0xEB);
                    HighMarkerColor.Color = Color.FromArgb(0xFF, 0x56, 0xBE, 0x79);
                    LowMarkerColor.Color = Color.FromArgb(0xFF, 0xFF, 0x50, 0x55);
                    break;
                case SLSparklineStyle.Colorful3:
                    SeriesColor.Color = Color.FromArgb(0xFF, 0xC6, 0xEF, 0xCE);
                    NegativeColor.Color = Color.FromArgb(0xFF, 0xFF, 0xC7, 0xCE);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.Color = Color.FromArgb(0xFF, 0x8C, 0xAD, 0xD6);
                    FirstMarkerColor.Color = Color.FromArgb(0xFF, 0xFF, 0xDC, 0x47);
                    LastMarkerColor.Color = Color.FromArgb(0xFF, 0xFF, 0xEB, 0x9C);
                    HighMarkerColor.Color = Color.FromArgb(0xFF, 0x60, 0xD2, 0x76);
                    LowMarkerColor.Color = Color.FromArgb(0xFF, 0xFF, 0x53, 0x67);
                    break;
                case SLSparklineStyle.Colorful4:
                    SeriesColor.Color = Color.FromArgb(0xFF, 0, 0xB0, 0x50);
                    NegativeColor.Color = Color.FromArgb(0xFF, 0xFF, 0, 0);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.Color = Color.FromArgb(0xFF, 0, 0x70, 0xC0);
                    FirstMarkerColor.Color = Color.FromArgb(0xFF, 0xFF, 0xC0, 0);
                    LastMarkerColor.Color = Color.FromArgb(0xFF, 0xFF, 0xC0, 0);
                    HighMarkerColor.Color = Color.FromArgb(0xFF, 0, 0xB0, 0x50);
                    LowMarkerColor.Color = Color.FromArgb(0xFF, 0xFF, 0, 0);
                    break;
                case SLSparklineStyle.Colorful5:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Dark2Color);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color);
                    break;
                case SLSparklineStyle.Colorful6:
                    SeriesColor.SetThemeColor(SLThemeColorIndexValues.Dark1Color);
                    NegativeColor.SetThemeColor(SLThemeColorIndexValues.Accent6Color);
                    AxisColor.Color = Color.FromArgb(0xFF, 0, 0, 0);
                    MarkersColor.SetThemeColor(SLThemeColorIndexValues.Accent5Color);
                    FirstMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent1Color);
                    LastMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent2Color);
                    HighMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent3Color);
                    LowMarkerColor.SetThemeColor(SLThemeColorIndexValues.Accent4Color);
                    break;
            }
        }

        internal void FromSparklineGroup(X14.SparklineGroup spkgrp)
        {
            SetAllNull();

            if (spkgrp.SeriesColor != null) SeriesColor.FromSeriesColor(spkgrp.SeriesColor);
            if (spkgrp.NegativeColor != null) NegativeColor.FromNegativeColor(spkgrp.NegativeColor);
            if (spkgrp.AxisColor != null) AxisColor.FromAxisColor(spkgrp.AxisColor);
            if (spkgrp.MarkersColor != null) MarkersColor.FromMarkersColor(spkgrp.MarkersColor);
            if (spkgrp.FirstMarkerColor != null) FirstMarkerColor.FromFirstMarkerColor(spkgrp.FirstMarkerColor);
            if (spkgrp.LastMarkerColor != null) LastMarkerColor.FromLastMarkerColor(spkgrp.LastMarkerColor);
            if (spkgrp.HighMarkerColor != null) HighMarkerColor.FromHighMarkerColor(spkgrp.HighMarkerColor);
            if (spkgrp.LowMarkerColor != null) LowMarkerColor.FromLowMarkerColor(spkgrp.LowMarkerColor);

            int index;
            var sRef = string.Empty;
            var sWorksheetName = string.Empty;
            var iStartRowIndex = -1;
            var iStartColumnIndex = -1;
            var iEndRowIndex = -1;
            var iEndColumnIndex = -1;

            if (spkgrp.Formula != null)
            {
                sRef = spkgrp.Formula.Text;
                index = sRef.IndexOf("!");
                if (index >= 0)
                {
                    DateWorksheetName = sRef.Substring(0, index);
                    sRef = sRef.Substring(index + 1);
                }

                index = sRef.LastIndexOf(":");

                if (index >= 0)
                {
                    if (
                        !SLTool.FormatCellReferenceRangeToRowColumnIndex(sRef, out iStartRowIndex, out iStartColumnIndex,
                            out iEndRowIndex, out iEndColumnIndex))
                    {
                        iStartRowIndex = -1;
                        iStartColumnIndex = -1;
                        iEndRowIndex = -1;
                        iEndColumnIndex = -1;
                    }

                    if ((iStartRowIndex > 0) && (iStartColumnIndex > 0) && (iEndRowIndex > 0) && (iEndColumnIndex > 0))
                    {
                        DateStartRowIndex = iStartRowIndex;
                        DateStartColumnIndex = iStartColumnIndex;
                        DateEndRowIndex = iEndRowIndex;
                        DateEndColumnIndex = iEndColumnIndex;
                        DateAxis = true;
                    }
                    else
                    {
                        DateStartRowIndex = 1;
                        DateStartColumnIndex = 1;
                        DateEndRowIndex = 1;
                        DateEndColumnIndex = 1;
                        DateAxis = false;
                    }
                }
                else
                {
                    if (!SLTool.FormatCellReferenceToRowColumnIndex(sRef, out iStartRowIndex, out iStartColumnIndex))
                    {
                        iStartRowIndex = -1;
                        iStartColumnIndex = -1;
                    }

                    if ((iStartRowIndex > 0) && (iStartColumnIndex > 0))
                    {
                        DateStartRowIndex = iStartRowIndex;
                        DateStartColumnIndex = iStartColumnIndex;
                        DateEndRowIndex = iStartRowIndex;
                        DateEndColumnIndex = iStartColumnIndex;
                        DateAxis = true;
                    }
                    else
                    {
                        DateStartRowIndex = 1;
                        DateStartColumnIndex = 1;
                        DateEndRowIndex = 1;
                        DateEndColumnIndex = 1;
                        DateAxis = false;
                    }
                }
            }

            if (spkgrp.Sparklines != null)
            {
                X14.Sparkline spkline;
                SLSparkline spk;
                foreach (var child in spkgrp.Sparklines.ChildElements)
                    if (child is X14.Sparkline)
                    {
                        spkline = (X14.Sparkline) child;
                        spk = new SLSparkline();
                        // the formula part contains the data source. Apparently, Excel is fine
                        // if it's empty. IF IT'S EMPTY THEN DELETE THE WHOLE SPARKLINE!
                        // Ok, I'm fine now... I'm gonna treat an empty Formula as "invalid".
                        if ((spkline.Formula != null) && (spkline.ReferenceSequence != null))
                        {
                            sRef = spkline.Formula.Text;
                            index = sRef.IndexOf("!");
                            if (index >= 0)
                            {
                                spk.WorksheetName = sRef.Substring(0, index);
                                sRef = sRef.Substring(index + 1);
                            }

                            index = sRef.LastIndexOf(":");

                            if (index >= 0)
                            {
                                if (
                                    !SLTool.FormatCellReferenceRangeToRowColumnIndex(sRef, out iStartRowIndex,
                                        out iStartColumnIndex, out iEndRowIndex, out iEndColumnIndex))
                                {
                                    iStartRowIndex = -1;
                                    iStartColumnIndex = -1;
                                    iEndRowIndex = -1;
                                    iEndColumnIndex = -1;
                                }

                                if ((iStartRowIndex > 0) && (iStartColumnIndex > 0) && (iEndRowIndex > 0) &&
                                    (iEndColumnIndex > 0))
                                {
                                    spk.StartRowIndex = iStartRowIndex;
                                    spk.StartColumnIndex = iStartColumnIndex;
                                    spk.EndRowIndex = iEndRowIndex;
                                    spk.EndColumnIndex = iEndColumnIndex;
                                }
                                else
                                {
                                    spk.StartRowIndex = 1;
                                    spk.StartColumnIndex = 1;
                                    spk.EndRowIndex = 1;
                                    spk.EndColumnIndex = 1;
                                }
                            }
                            else
                            {
                                if (
                                    !SLTool.FormatCellReferenceToRowColumnIndex(sRef, out iStartRowIndex,
                                        out iStartColumnIndex))
                                {
                                    iStartRowIndex = -1;
                                    iStartColumnIndex = -1;
                                }

                                if ((iStartRowIndex > 0) && (iStartColumnIndex > 0))
                                {
                                    spk.StartRowIndex = iStartRowIndex;
                                    spk.StartColumnIndex = iStartColumnIndex;
                                    spk.EndRowIndex = iStartRowIndex;
                                    spk.EndColumnIndex = iStartColumnIndex;
                                }
                                else
                                {
                                    spk.StartRowIndex = 1;
                                    spk.StartColumnIndex = 1;
                                    spk.EndRowIndex = 1;
                                    spk.EndColumnIndex = 1;
                                }
                            }

                            if (
                                !SLTool.FormatCellReferenceToRowColumnIndex(spkline.ReferenceSequence.Text,
                                    out iStartRowIndex, out iStartColumnIndex))
                            {
                                iStartRowIndex = -1;
                                iStartColumnIndex = -1;
                            }

                            if ((iStartRowIndex > 0) && (iStartColumnIndex > 0))
                            {
                                spk.LocationRowIndex = iStartRowIndex;
                                spk.LocationColumnIndex = iStartColumnIndex;

                                // there are so many things that could possibly go wrong
                                // that we'll just assume that if the location part is correct,
                                // we'll just take it...
                                Sparklines.Add(spk.Clone());
                            }
                        }
                    }
            }

            if (spkgrp.ManualMax != null) ManualMax = spkgrp.ManualMax.Value;
            if (spkgrp.ManualMin != null) ManualMin = spkgrp.ManualMin.Value;
            if (spkgrp.LineWeight != null) LineWeight = (decimal) spkgrp.LineWeight.Value;
            if (spkgrp.Type != null) Type = spkgrp.Type.Value;

            // we're gonna ignore dateAxis because if there's no formula, having it true is useless

            if (spkgrp.DisplayEmptyCellsAs != null) ShowEmptyCellsAs = spkgrp.DisplayEmptyCellsAs.Value;
            if (spkgrp.Markers != null) ShowMarkers = spkgrp.Markers.Value;
            if (spkgrp.High != null) ShowHighPoint = spkgrp.High.Value;
            if (spkgrp.Low != null) ShowLowPoint = spkgrp.Low.Value;
            if (spkgrp.First != null) ShowFirstPoint = spkgrp.First.Value;
            if (spkgrp.Last != null) ShowLastPoint = spkgrp.Last.Value;
            if (spkgrp.Negative != null) ShowNegativePoints = spkgrp.Negative.Value;
            if (spkgrp.DisplayXAxis != null) ShowAxis = spkgrp.DisplayXAxis.Value;
            if (spkgrp.DisplayHidden != null) ShowHiddenData = spkgrp.DisplayHidden.Value;
            if (spkgrp.MinAxisType != null) MinAxisType = spkgrp.MinAxisType.Value;
            if (spkgrp.MaxAxisType != null) MaxAxisType = spkgrp.MaxAxisType.Value;
            if (spkgrp.RightToLeft != null) RightToLeft = spkgrp.RightToLeft.Value;
        }

        internal X14.SparklineGroup ToSparklineGroup()
        {
            var spkgrp = new X14.SparklineGroup();

            if (!SeriesColor.IsEmpty()) spkgrp.SeriesColor = SeriesColor.ToSeriesColor();
            if (!NegativeColor.IsEmpty()) spkgrp.NegativeColor = NegativeColor.ToNegativeColor();
            if (!AxisColor.IsEmpty()) spkgrp.AxisColor = AxisColor.ToAxisColor();
            if (!MarkersColor.IsEmpty()) spkgrp.MarkersColor = MarkersColor.ToMarkersColor();
            if (!FirstMarkerColor.IsEmpty()) spkgrp.FirstMarkerColor = FirstMarkerColor.ToFirstMarkerColor();
            if (!LastMarkerColor.IsEmpty()) spkgrp.LastMarkerColor = LastMarkerColor.ToLastMarkerColor();
            if (!HighMarkerColor.IsEmpty()) spkgrp.HighMarkerColor = HighMarkerColor.ToHighMarkerColor();
            if (!LowMarkerColor.IsEmpty()) spkgrp.LowMarkerColor = LowMarkerColor.ToLowMarkerColor();

            if (DateAxis)
            {
                if ((DateStartRowIndex == DateEndRowIndex) && (DateStartColumnIndex == DateEndColumnIndex))
                {
                    spkgrp.Formula = new Formula();
                    spkgrp.Formula.Text = SLTool.ToCellReference(DateWorksheetName, DateStartRowIndex,
                        DateStartColumnIndex);
                }
                else
                {
                    spkgrp.Formula = new Formula();
                    spkgrp.Formula.Text = SLTool.ToCellRange(DateWorksheetName, DateStartRowIndex, DateStartColumnIndex,
                        DateEndRowIndex, DateEndColumnIndex);
                }

                spkgrp.DateAxis = true;
            }

            spkgrp.Sparklines = new X14.Sparklines();
            foreach (var spk in Sparklines)
                spkgrp.Sparklines.Append(spk.ToSparkline());

            switch (MinAxisType)
            {
                case X14.SparklineAxisMinMaxValues.Individual:
                    // default, so don't have to do anything
                    break;
                case X14.SparklineAxisMinMaxValues.Group:
                    spkgrp.MinAxisType = X14.SparklineAxisMinMaxValues.Group;
                    break;
                case X14.SparklineAxisMinMaxValues.Custom:
                    spkgrp.MinAxisType = X14.SparklineAxisMinMaxValues.Custom;
                    spkgrp.ManualMin = ManualMin;
                    break;
            }

            switch (MaxAxisType)
            {
                case X14.SparklineAxisMinMaxValues.Individual:
                    // default, so don't have to do anything
                    break;
                case X14.SparklineAxisMinMaxValues.Group:
                    spkgrp.MaxAxisType = X14.SparklineAxisMinMaxValues.Group;
                    break;
                case X14.SparklineAxisMinMaxValues.Custom:
                    spkgrp.MaxAxisType = X14.SparklineAxisMinMaxValues.Custom;
                    spkgrp.ManualMax = ManualMax;
                    break;
            }

            if (decLineWeight != 0.75m) spkgrp.LineWeight = (double) decLineWeight;

            if (Type != X14.SparklineTypeValues.Line) spkgrp.Type = Type;

            if (ShowEmptyCellsAs != X14.DisplayBlanksAsValues.Zero) spkgrp.DisplayEmptyCellsAs = ShowEmptyCellsAs;

            if (ShowMarkers) spkgrp.Markers = true;
            if (ShowHighPoint) spkgrp.High = true;
            if (ShowLowPoint) spkgrp.Low = true;
            if (ShowFirstPoint) spkgrp.First = true;
            if (ShowLastPoint) spkgrp.Last = true;
            if (ShowNegativePoints) spkgrp.Negative = true;
            if (ShowAxis) spkgrp.DisplayXAxis = true;
            if (ShowHiddenData) spkgrp.DisplayHidden = true;
            if (RightToLeft) spkgrp.RightToLeft = true;

            return spkgrp;
        }

        internal SLSparklineGroup Clone()
        {
            var spkgrp = new SLSparklineGroup(listThemeColors, listIndexedColors);
            spkgrp.WorksheetName = WorksheetName;
            spkgrp.StartRowIndex = StartRowIndex;
            spkgrp.StartColumnIndex = StartColumnIndex;
            spkgrp.EndRowIndex = EndRowIndex;
            spkgrp.EndColumnIndex = EndColumnIndex;

            spkgrp.SeriesColor = SeriesColor.Clone();
            spkgrp.NegativeColor = NegativeColor.Clone();
            spkgrp.AxisColor = AxisColor.Clone();
            spkgrp.MarkersColor = MarkersColor.Clone();
            spkgrp.FirstMarkerColor = FirstMarkerColor.Clone();
            spkgrp.LastMarkerColor = LastMarkerColor.Clone();
            spkgrp.HighMarkerColor = HighMarkerColor.Clone();
            spkgrp.LowMarkerColor = LowMarkerColor.Clone();

            spkgrp.DateWorksheetName = DateWorksheetName;
            spkgrp.DateStartRowIndex = DateStartRowIndex;
            spkgrp.DateStartColumnIndex = DateStartColumnIndex;
            spkgrp.DateEndRowIndex = DateEndRowIndex;
            spkgrp.DateEndColumnIndex = DateEndColumnIndex;
            spkgrp.DateAxis = DateAxis;

            foreach (var spk in Sparklines)
                spkgrp.Sparklines.Add(spk.Clone());

            spkgrp.ManualMax = ManualMax;
            spkgrp.MaxAxisType = MaxAxisType;
            spkgrp.ManualMin = ManualMin;
            spkgrp.MinAxisType = MinAxisType;

            spkgrp.decLineWeight = decLineWeight;

            spkgrp.Type = Type;

            spkgrp.ShowEmptyCellsAs = ShowEmptyCellsAs;

            spkgrp.ShowMarkers = ShowMarkers;
            spkgrp.ShowHighPoint = ShowHighPoint;
            spkgrp.ShowLowPoint = ShowLowPoint;
            spkgrp.ShowFirstPoint = ShowFirstPoint;
            spkgrp.ShowLastPoint = ShowLastPoint;
            spkgrp.ShowNegativePoints = ShowNegativePoints;
            spkgrp.ShowAxis = ShowAxis;
            spkgrp.ShowHiddenData = ShowHiddenData;
            spkgrp.RightToLeft = RightToLeft;

            return spkgrp;
        }
    }
}