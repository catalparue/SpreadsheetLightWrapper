using System.Collections.Generic;
using System.Drawing;

namespace SpreadsheetLightWrapper.Core.Charts
{
    /// <summary>
    ///     Chart customization options for area charts.
    /// </summary>
    public class SLAreaChartOptions
    {
        internal ushort iGapDepth;

        /// <summary>
        ///     Initializes an instance of SLAreaChartOptions. It is recommended to use SLChart.CreateAreaChartOptions().
        /// </summary>
        public SLAreaChartOptions()
        {
            Initialize(new List<Color>(), false);
        }

        internal SLAreaChartOptions(List<Color> ThemeColors, bool IsStylish = false)
        {
            Initialize(ThemeColors, IsStylish);
        }

        /// <summary>
        ///     Indicates if the area chart has drop lines.
        /// </summary>
        public bool HasDropLines { get; set; }

        /// <summary>
        ///     Drop lines properties.
        /// </summary>
        public SLDropLines DropLines { get; set; }

        /// <summary>
        ///     The gap depth between area clusters (between different data series) as a percentage of width, ranging between 0%
        ///     and 500% (both inclusive). The default is 150%. This is only used for 3D chart version.
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

        private void Initialize(List<Color> ThemeColors, bool IsStylish)
        {
            HasDropLines = false;
            DropLines = new SLDropLines(ThemeColors, IsStylish);
            iGapDepth = 150;
        }

        /// <summary>
        ///     Clone a new instance of SLAreaChartOptions.
        /// </summary>
        /// <returns>An SLAreaChartOptions object.</returns>
        public SLAreaChartOptions Clone()
        {
            var aco = new SLAreaChartOptions();
            aco.HasDropLines = HasDropLines;
            aco.DropLines = DropLines.Clone();
            aco.iGapDepth = iGapDepth;

            return aco;
        }
    }
}