using System.Collections.Generic;
using System.Drawing;
using Ups.Toolkit.SpreadsheetLight.Core.misc;
using A = DocumentFormat.OpenXml.Drawing;

namespace Ups.Toolkit.SpreadsheetLight.Core.Drawing
{
    /// <summary>
    ///     Encapsulates properties and methods for specifying glow effects.
    ///     This simulates the DocumentFormat.OpenXml.Drawing.Glow class.
    /// </summary>
    public class SLGlow
    {
        private decimal decRadius;
        internal bool HasGlow;

        internal SLGlow(List<Color> ThemeColors)
        {
            HasGlow = false;
            decRadius = 0;
            GlowColor = new SLColorTransform(ThemeColors);
        }

        internal decimal Radius
        {
            get { return decRadius; }
            set
            {
                decRadius = value;
                if (decRadius < 0m) decRadius = 0m;
                if (decRadius > 2147483647m) decRadius = 2147483647m;
            }
        }

        internal SLColorTransform GlowColor { get; set; }

        /// <summary>
        ///     Set no glow.
        /// </summary>
        public void SetNoGlow()
        {
            HasGlow = false;
            decRadius = 0;
        }

        /// <summary>
        ///     Set the glow color.
        /// </summary>
        /// <param name="GlowColor">The color used.</param>
        /// <param name="Transparency">Transparency of the color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="Size">The size in points. The suggested range is 0 pt to 150 pt (both inclusive).</param>
        public void SetGlow(Color GlowColor, decimal Transparency, decimal Size)
        {
            HasGlow = true;
            this.GlowColor.SetColor(GlowColor, Transparency);
            Radius = Size;
        }

        /// <summary>
        ///     Set the glow color.
        /// </summary>
        /// <param name="GlowColor">The theme color used.</param>
        /// <param name="Tint">
        ///     The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color
        ///     and positive tints lighten the theme color.
        /// </param>
        /// <param name="Transparency">Transparency of the color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="Size">The size in points. The suggested range is 0 pt to 150 pt (both inclusive).</param>
        public void SetGlow(SLThemeColorIndexValues GlowColor, double Tint, decimal Transparency, decimal Size)
        {
            HasGlow = true;
            this.GlowColor.SetColor(GlowColor, Tint, Transparency);
            Radius = Size;
        }

        internal A.Glow ToGlow()
        {
            var g = new A.Glow();
            if (GlowColor.IsRgbColorModelHex)
                g.RgbColorModelHex = GlowColor.ToRgbColorModelHex();
            else
                g.SchemeColor = GlowColor.ToSchemeColor();
            g.Radius = SLDrawingTool.CalculatePositiveCoordinate(decRadius);

            return g;
        }

        internal SLGlow Clone()
        {
            var g = new SLGlow(GlowColor.listThemeColors);
            g.HasGlow = HasGlow;
            g.decRadius = decRadius;
            g.GlowColor = GlowColor.Clone();

            return g;
        }
    }
}