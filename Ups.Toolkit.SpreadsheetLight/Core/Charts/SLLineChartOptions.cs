using System.Collections.Generic;
using System.Drawing;

namespace Ups.Toolkit.SpreadsheetLight.Core.Charts
{
    /// <summary>
    ///     Chart customization options for line charts.
    /// </summary>
    public class SLLineChartOptions
    {
        internal ushort iGapDepth;

        /// <summary>
        ///     Initializes an instance of SLLineChartOptions. It is recommended to use SLChart.CreateLineChartOptions().
        /// </summary>
        public SLLineChartOptions()
        {
            Initialize(new List<Color>(), false);
        }

        internal SLLineChartOptions(List<Color> ThemeColors, bool IsStylish = false)
        {
            Initialize(ThemeColors, IsStylish);
        }

        /// <summary>
        ///     The gap depth between line clusters (between different data series) as a percentage of bar or column width, ranging
        ///     between 0% and 500% (both inclusive). The default is 150%. This is only used for 3D chart version.
        /// </summary>
        public ushort GapDepth
        {
            get { return iGapDepth; }
            set
            {
                iGapDepth = value;
                if (iGapDepth > 500) iGapDepth = 500;
            }
        }

        /// <summary>
        ///     Indicates if the line chart has drop lines.
        /// </summary>
        public bool HasDropLines { get; set; }

        /// <summary>
        ///     Drop lines properties.
        /// </summary>
        public SLDropLines DropLines { get; set; }

        /// <summary>
        ///     Indicates if the line chart has high-low lines. This is not applicable for 3D line charts.
        /// </summary>
        public bool HasHighLowLines { get; set; }

        /// <summary>
        ///     High-low lines properties.
        /// </summary>
        public SLHighLowLines HighLowLines { get; set; }

        /// <summary>
        ///     Indicates if the line chart has up-down bars. This is not applicable for 3D line charts.
        /// </summary>
        public bool HasUpDownBars { get; set; }

        /// <summary>
        ///     Up-down bars properties.
        /// </summary>
        public SLUpDownBars UpDownBars { get; set; }

        /// <summary>
        ///     Whether the line connecting data points use C splines (instead of straight lines).
        /// </summary>
        public bool Smooth { get; set; }

        private void Initialize(List<Color> ThemeColors, bool IsStylish)
        {
            iGapDepth = 150;
            HasDropLines = false;
            DropLines = new SLDropLines(ThemeColors, IsStylish);
            HasHighLowLines = false;
            HighLowLines = new SLHighLowLines(ThemeColors, IsStylish);
            HasUpDownBars = false;
            UpDownBars = new SLUpDownBars(ThemeColors, IsStylish);
            Smooth = false;
        }

        /// <summary>
        ///     Clone an instance of SLLineChartOptions.
        /// </summary>
        /// <returns>An SLLineChartOptions object.</returns>
        public SLLineChartOptions Clone()
        {
            var lco = new SLLineChartOptions();
            lco.iGapDepth = iGapDepth;
            lco.HasDropLines = HasDropLines;
            lco.DropLines = DropLines.Clone();
            lco.HasHighLowLines = HasHighLowLines;
            lco.HighLowLines = HighLowLines.Clone();
            lco.HasUpDownBars = HasUpDownBars;
            lco.UpDownBars = UpDownBars.Clone();
            lco.Smooth = Smooth;

            return lco;
        }
    }
}