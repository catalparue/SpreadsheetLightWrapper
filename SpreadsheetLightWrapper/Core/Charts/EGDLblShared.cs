using System.Collections.Generic;
using System.Drawing;
using SpreadsheetLightWrapper.Core.Drawing;
using SpreadsheetLightWrapper.Core.misc;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLightWrapper.Core.Charts
{
    /// <summary>
    ///     This simulates the element group EG_DLblShared as specified in the Open XML specs.
    /// </summary>
    public abstract class EGDLblShared : SLChartAlignment
    {
        internal bool bSourceLinked;
        // This is C.NumberingFormat
        internal bool HasNumberingFormat;

        internal string sFormatCode;

        internal C.DataLabelPositionValues? vLabelPosition;

        internal EGDLblShared(List<Color> ThemeColors)
        {
            sFormatCode = SLConstants.NumberFormatGeneral;
            bSourceLinked = true;
            HasNumberingFormat = false;
            vLabelPosition = null;
            ShapeProperties = new SLShapeProperties(ThemeColors);
            ShowLegendKey = false;
            ShowValue = false;
            ShowCategoryName = false;
            ShowSeriesName = false;
            ShowPercentage = false;
            ShowBubbleSize = false;
            Separator = string.Empty;
        }

        /// <summary>
        ///     Format code. If you set a custom format code, you might also want to set SourceLinked to false.
        /// </summary>
        public string FormatCode
        {
            get { return sFormatCode; }
            set
            {
                sFormatCode = value;
                HasNumberingFormat = true;
            }
        }

        /// <summary>
        ///     Whether the format code is linked to the data source.
        /// </summary>
        public bool SourceLinked
        {
            get { return bSourceLinked; }
            set
            {
                bSourceLinked = value;
                HasNumberingFormat = true;
            }
        }

        internal SLShapeProperties ShapeProperties { get; set; }

        /// <summary>
        ///     Fill properties.
        /// </summary>
        public SLFill Fill
        {
            get { return ShapeProperties.Fill; }
        }

        /// <summary>
        ///     Border properties.
        /// </summary>
        public SLLinePropertiesType Border
        {
            get { return ShapeProperties.Outline; }
        }

        /// <summary>
        ///     Shadow properties.
        /// </summary>
        public SLShadowEffect Shadow
        {
            get { return ShapeProperties.EffectList.Shadow; }
        }

        /// <summary>
        ///     Glow properties.
        /// </summary>
        public SLGlow Glow
        {
            get { return ShapeProperties.EffectList.Glow; }
        }

        /// <summary>
        ///     Soft edge properties.
        /// </summary>
        public SLSoftEdge SoftEdge
        {
            get { return ShapeProperties.EffectList.SoftEdge; }
        }

        /// <summary>
        ///     3D format properties.
        /// </summary>
        public SLFormat3D Format3D
        {
            get { return ShapeProperties.Format3D; }
        }

        /// <summary>
        ///     Specifies if the legend key is included in the label.
        /// </summary>
        public bool ShowLegendKey { get; set; }

        /// <summary>
        ///     Specifies if the label contains the value. For certain charts, this is known as the "Y Value".
        /// </summary>
        public bool ShowValue { get; set; }

        /// <summary>
        ///     Specifies if the label contains the category name. For certain charts, this is known as the "X Value".
        /// </summary>
        public bool ShowCategoryName { get; set; }

        /// <summary>
        ///     Specifies if the label contains the series name.
        /// </summary>
        public bool ShowSeriesName { get; set; }

        /// <summary>
        ///     Specifies if the label contains the percentage. This is for pie charts.
        /// </summary>
        public bool ShowPercentage { get; set; }

        /// <summary>
        ///     Specifies if the label contains the bubble size. This is for bubble charts.
        /// </summary>
        public bool ShowBubbleSize { get; set; }

        /// <summary>
        ///     The separator.
        /// </summary>
        public string Separator { get; set; }

        /// <summary>
        ///     Set the position of the data label.
        /// </summary>
        /// <param name="Position">The data label position.</param>
        public void SetLabelPosition(C.DataLabelPositionValues Position)
        {
            vLabelPosition = Position;
        }

        /// <summary>
        ///     Set automatic positioning of the data label.
        /// </summary>
        public void SetAutoLabelPosition()
        {
            vLabelPosition = null;
        }
    }
}