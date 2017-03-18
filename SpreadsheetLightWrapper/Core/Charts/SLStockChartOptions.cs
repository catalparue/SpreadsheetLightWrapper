using System.Collections.Generic;
using System.Drawing;
using SpreadsheetLightWrapper.Core.Drawing;

namespace SpreadsheetLightWrapper.Core.Charts
{
    /// <summary>
    ///     Chart customization options for stock charts.
    /// </summary>
    public class SLStockChartOptions
    {
        internal sbyte byOverlap;
        internal ushort iGapWidth;

        /// <summary>
        ///     Initializes an instance of SLStockChartOptions. It is recommended to use SLChart.CreateStockChartOptions().
        /// </summary>
        public SLStockChartOptions()
        {
            Initialize(new List<Color>(), false);
        }

        internal SLStockChartOptions(List<Color> ThemeColors, bool IsStylish = false)
        {
            Initialize(ThemeColors, IsStylish);
        }

        /// <summary>
        ///     The gap width between columns as a percentage of column width, ranging between 0% and 500% (both inclusive). The
        ///     default is 150%.
        ///     This only applies when there's Volume data.
        /// </summary>
        public ushort GapWidth
        {
            get { return iGapWidth; }
            set
            {
                iGapWidth = value;
                if (iGapWidth > 500) iGapWidth = 500;
            }
        }

        /// <summary>
        ///     The amount of overlapping for columns, ranging from -100 to 100 (both inclusive). The default is 0.
        ///     This only applies when there's Volume data.
        /// </summary>
        public sbyte Overlap
        {
            get { return byOverlap; }
            set
            {
                byOverlap = value;
                if (byOverlap < -100) byOverlap = -100;
                if (byOverlap > 100) byOverlap = 100;
            }
        }

        internal SLShapeProperties ShapeProperties { get; set; }

        /// <summary>
        ///     Fill properties for Volume data.
        /// </summary>
        public SLFill Fill
        {
            get { return ShapeProperties.Fill; }
        }

        /// <summary>
        ///     Border properties for Volume data.
        /// </summary>
        public SLLinePropertiesType Border
        {
            get { return ShapeProperties.Outline; }
        }

        /// <summary>
        ///     Shadow properties for Volume data.
        /// </summary>
        public SLShadowEffect Shadow
        {
            get { return ShapeProperties.EffectList.Shadow; }
        }

        /// <summary>
        ///     Glow properties for Volume data.
        /// </summary>
        public SLGlow Glow
        {
            get { return ShapeProperties.EffectList.Glow; }
        }

        /// <summary>
        ///     Soft edge properties for Volume data.
        /// </summary>
        public SLSoftEdge SoftEdge
        {
            get { return ShapeProperties.EffectList.SoftEdge; }
        }

        /// <summary>
        ///     3D format properties for Volume data.
        /// </summary>
        public SLFormat3D Format3D
        {
            get { return ShapeProperties.Format3D; }
        }

        /// <summary>
        ///     Indicates if the stock chart has drop lines.
        /// </summary>
        public bool HasDropLines { get; set; }

        /// <summary>
        ///     Drop lines properties.
        /// </summary>
        public SLDropLines DropLines { get; set; }

        /// <summary>
        ///     Indicates if the stock chart has high-low lines.
        /// </summary>
        public bool HasHighLowLines { get; set; }

        /// <summary>
        ///     High-low lines properties.
        /// </summary>
        public SLHighLowLines HighLowLines { get; set; }

        /// <summary>
        ///     Indicates if the stock chart has up-down bars.
        /// </summary>
        public bool HasUpDownBars { get; set; }

        /// <summary>
        ///     Up-down bars properties.
        /// </summary>
        public SLUpDownBars UpDownBars { get; set; }

        private void Initialize(List<Color> ThemeColors, bool IsStylish)
        {
            iGapWidth = 150;
            byOverlap = 0;
            ShapeProperties = new SLShapeProperties(ThemeColors);
            HasDropLines = false;
            DropLines = new SLDropLines(ThemeColors, IsStylish);
            HasHighLowLines = true;
            HighLowLines = new SLHighLowLines(ThemeColors, IsStylish);
            HasUpDownBars = true;
            UpDownBars = new SLUpDownBars(ThemeColors, IsStylish);
        }

        internal SLStockChartOptions Clone()
        {
            var sco = new SLStockChartOptions();
            sco.iGapWidth = iGapWidth;
            sco.byOverlap = byOverlap;
            sco.ShapeProperties = ShapeProperties.Clone();
            sco.HasDropLines = HasDropLines;
            sco.DropLines = DropLines.Clone();
            sco.HasHighLowLines = HasHighLowLines;
            sco.HighLowLines = HighLowLines.Clone();
            sco.HasUpDownBars = HasUpDownBars;
            sco.UpDownBars = UpDownBars.Clone();

            return sco;
        }
    }
}