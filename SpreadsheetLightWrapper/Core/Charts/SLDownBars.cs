using System.Collections.Generic;
using System.Drawing;
using SpreadsheetLightWrapper.Core.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLightWrapper.Core.Charts
{
    /// <summary>
    ///     Encapsulates properties and methods for down bars.
    ///     This simulates the DocumentFormat.OpenXml.Drawing.Charts.DownBars class.
    /// </summary>
    public class SLDownBars
    {
        internal SLDownBars(List<Color> ThemeColors, bool IsStylish = false)
        {
            ShapeProperties = new SLShapeProperties(ThemeColors);
            if (IsStylish)
            {
                ShapeProperties.Fill.SetSolidFill(A.SchemeColorValues.Dark1, 0.35m, 0);
                ShapeProperties.Outline.Width = 0.75m;
                ShapeProperties.Outline.SetSolidLine(A.SchemeColorValues.Text1, 0.35m, 0);
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

        internal C.DownBars ToDownBars(bool IsStylish = false)
        {
            var db = new C.DownBars();

            if (ShapeProperties.HasShapeProperties)
                db.ChartShapeProperties = ShapeProperties.ToChartShapeProperties(IsStylish);

            return db;
        }

        internal SLDownBars Clone()
        {
            var db = new SLDownBars(ShapeProperties.listThemeColors);
            db.ShapeProperties = ShapeProperties.Clone();

            return db;
        }
    }
}