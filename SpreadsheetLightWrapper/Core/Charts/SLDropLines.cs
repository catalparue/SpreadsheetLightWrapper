using System.Collections.Generic;
using System.Drawing;
using SpreadsheetLightWrapper.Core.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLightWrapper.Core.Charts
{
    /// <summary>
    ///     Encapsulates properties and methods for drop lines.
    ///     This simulates the DocumentFormat.OpenXml.Drawing.Charts.DropLines class.
    /// </summary>
    public class SLDropLines
    {
        internal SLDropLines(List<Color> ThemeColors, bool IsStylish = false)
        {
            ShapeProperties = new SLShapeProperties(ThemeColors);
            if (IsStylish)
            {
                ShapeProperties.Outline.Width = 0.75m;
                ShapeProperties.Outline.CapType = A.LineCapValues.Flat;
                ShapeProperties.Outline.CompoundLineType = A.CompoundLineValues.Single;
                ShapeProperties.Outline.Alignment = A.PenAlignmentValues.Center;
                ShapeProperties.Outline.SetSolidLine(A.SchemeColorValues.Text1, 0.65m, 0);
                ShapeProperties.Outline.JoinType = SLLineJoinValues.Round;
            }
        }

        internal SLShapeProperties ShapeProperties { get; set; }

        /// <summary>
        ///     Line properties.
        /// </summary>
        public SLLinePropertiesType Line
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
        ///     Clear all styling shape properties. Use this if you want to start styling from a clean slate.
        /// </summary>
        public void ClearShapeProperties()
        {
            ShapeProperties = new SLShapeProperties(ShapeProperties.listThemeColors);
        }

        internal C.DropLines ToDropLines(bool IsStylish = false)
        {
            var dl = new C.DropLines();
            dl.ChartShapeProperties = ShapeProperties.ToChartShapeProperties(IsStylish);

            return dl;
        }

        internal SLDropLines Clone()
        {
            var dl = new SLDropLines(ShapeProperties.listThemeColors);
            dl.ShapeProperties = ShapeProperties.Clone();

            return dl;
        }
    }
}