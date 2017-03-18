using System.Collections.Generic;
using System.Drawing;
using Ups.Toolkit.SpreadsheetLight.Core.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Ups.Toolkit.SpreadsheetLight.Core.Charts
{
    /// <summary>
    ///     Encapsulates properties and methods for up bars.
    ///     This simulates the DocumentFormat.OpenXml.Drawing.Charts.UpBars class.
    /// </summary>
    public class SLUpBars
    {
        internal SLUpBars(List<Color> ThemeColors, bool IsStylish = false)
        {
            ShapeProperties = new SLShapeProperties(ThemeColors);
            if (IsStylish)
            {
                ShapeProperties.Fill.SetSolidFill(A.SchemeColorValues.Light1, 0, 0);
                ShapeProperties.Outline.Width = 0.75m;
                ShapeProperties.Outline.SetSolidLine(A.SchemeColorValues.Text1, 0.85m, 0);
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
        ///     Clear all styling shape properties. Use this if you want to start styling from a clean slate.
        /// </summary>
        public void ClearShapeProperties()
        {
            ShapeProperties = new SLShapeProperties(ShapeProperties.listThemeColors);
        }

        internal C.UpBars ToUpBars(bool IsStylish = false)
        {
            var ub = new C.UpBars();

            if (ShapeProperties.HasShapeProperties)
                ub.ChartShapeProperties = ShapeProperties.ToChartShapeProperties(IsStylish);

            return ub;
        }

        internal SLUpBars Clone()
        {
            var ub = new SLUpBars(ShapeProperties.listThemeColors);
            ub.ShapeProperties = ShapeProperties.Clone();

            return ub;
        }
    }
}