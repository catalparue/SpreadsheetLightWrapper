using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using A = DocumentFormat.OpenXml.Drawing;

namespace SpreadsheetLightWrapper.Core.Drawing
{
    internal class SLGradientStop
    {
        private decimal decPosition;

        internal SLGradientStop(List<Color> ThemeColors)
        {
            Color = new SLColorTransform(ThemeColors);
            Position = 0m;
        }

        internal SLGradientStop(List<Color> ThemeColors, string HexColor, decimal Position)
        {
            Color = new SLColorTransform(ThemeColors);
            this.Position = Position;

            var clr = new Color();
            try
            {
                clr = System.Drawing.Color.FromArgb(int.Parse(HexColor, NumberStyles.HexNumber));
            }
            catch
            {
                clr = System.Drawing.Color.White;
            }
            Color.SetColor(clr, 0);
        }

        internal SLColorTransform Color { get; set; }

        /// <summary>
        ///     The position in percentage ranging from 0% to 100%. Accurate to 1/1000 of a percent.
        /// </summary>
        internal decimal Position
        {
            get { return decPosition; }
            set
            {
                decPosition = value;
                if (decPosition < 0m) decPosition = 0m;
                if (decPosition > 100m) decPosition = 100m;
            }
        }

        internal A.GradientStop ToGradientStop()
        {
            var gs = new A.GradientStop();
            if (Color.IsRgbColorModelHex) gs.RgbColorModelHex = Color.ToRgbColorModelHex();
            else gs.SchemeColor = Color.ToSchemeColor();

            gs.Position = Convert.ToInt32(Position*1000m);

            return gs;
        }

        internal SLGradientStop Clone()
        {
            var gs = new SLGradientStop(Color.listThemeColors);
            gs.Color = Color.Clone();
            gs.decPosition = decPosition;

            return gs;
        }
    }
}