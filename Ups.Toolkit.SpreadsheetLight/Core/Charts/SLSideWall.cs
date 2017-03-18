using System.Collections.Generic;
using System.Drawing;
using Ups.Toolkit.SpreadsheetLight.Core.Drawing;

namespace Ups.Toolkit.SpreadsheetLight.Core.Charts
{
    /// <summary>
    ///     Encapsulates properties and methods for setting the side wall of 3D charts.
    ///     This simulates the DocumentFormat.OpenXml.Drawing.Charts.SideWall class.
    /// </summary>
    public class SLSideWall
    {
        internal SLShapeProperties ShapeProperties;

        internal SLSideWall(List<Color> ThemeColors, bool IsStylish = false)
        {
            Thickness = 0;
            ShapeProperties = new SLShapeProperties(ThemeColors);
            if (IsStylish)
            {
                ShapeProperties.Fill.SetNoFill();
                ShapeProperties.Outline.SetNoLine();
            }
        }

        // From the Open XML SDK documentation:
        // "This element specifies the thickness of the walls or floor as a percentage of the largest dimension of the plot volume."
        // I have no idea what that means... and Excel doesn't allow the user to set this. Hmm...
        internal byte Thickness { get; set; }

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
        ///     3D rotation properties.
        /// </summary>
        public SLRotation3D Rotation3D
        {
            get { return ShapeProperties.Rotation3D; }
        }

        /// <summary>
        ///     Clear all styling shape properties. Use this if you want to start styling from a clean slate.
        /// </summary>
        public void ClearShapeProperties()
        {
            ShapeProperties = new SLShapeProperties(ShapeProperties.listThemeColors);
        }

        internal SLSideWall Clone()
        {
            var sw = new SLSideWall(ShapeProperties.listThemeColors);
            sw.Thickness = Thickness;
            sw.ShapeProperties = ShapeProperties.Clone();

            return sw;
        }
    }
}