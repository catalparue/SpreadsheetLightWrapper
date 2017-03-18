namespace Ups.Toolkit.SpreadsheetLight.Core.Charts
{
    /// <summary>
    ///     General chart customization options on chart creation.
    /// </summary>
    public class SLCreateChartOptions
    {
        /// <summary>
        ///     Initializes an instance of SLCreateChartOptions.
        /// </summary>
        public SLCreateChartOptions()
        {
            RowsAsDataSeries = null;
            ShowHiddenData = false;
            IsStylish = true;
        }

        /// <summary>
        ///     True if rows contain the data series. False if columns contain the data series.
        ///     Set to null if SpreadsheetLight is to determine data series orientation.
        ///     If the number of columns in a given cell range is more than or equal to the
        ///     number of rows, then it's decided that rows contain data series (else it's columns).
        ///     The default value is null.
        /// </summary>
        public bool? RowsAsDataSeries { get; set; }

        /// <summary>
        ///     True if hidden data is used in the chart. False otherwise.
        ///     The default value is false.
        /// </summary>
        public bool ShowHiddenData { get; set; }

        /// <summary>
        ///     True to use default chart styling from latest version of Excel
        ///     (but no guarantees on conformity or Excel version). False otherwise.
        ///     The default value is true.
        /// </summary>
        public bool IsStylish { get; set; }
    }
}