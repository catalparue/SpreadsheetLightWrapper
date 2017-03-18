using System;
using System.Collections.Generic;
using System.Drawing;
using SpreadsheetLightWrapper.Core.misc;
using A = DocumentFormat.OpenXml.Drawing;

namespace SpreadsheetLightWrapper.Core.Drawing
{
    internal class SLColorTransform
    {
        private decimal decTint;

        private decimal decTransparency;

        internal bool IsRgbColorModelHex;
        internal List<Color> listThemeColors;

        internal SLColorTransform(List<Color> ThemeColors)
        {
            int i;
            listThemeColors = new List<Color>();
            for (i = 0; i < ThemeColors.Count; ++i)
                listThemeColors.Add(ThemeColors[i]);

            SetAllNull();
        }

        /// <summary>
        ///     This is read-only
        /// </summary>
        internal Color DisplayColor { get; private set; }

        private Color RgbColor { get; set; }
        private A.SchemeColorValues SchemeColor { get; set; }

        private decimal Tint
        {
            get { return decTint; }
            set
            {
                decTint = value;
                if (decTint < -1.0m) decTint = -1.0m;
                if (decTint > 1.0m) decTint = 1.0m;
            }
        }

        internal decimal Transparency
        {
            get { return decTransparency; }
            set
            {
                decTransparency = value;
                if (decTransparency < 0m) decTransparency = 0m;
                if (decTransparency > 100m) decTransparency = 100m;
            }
        }

        private void SetAllNull()
        {
            IsRgbColorModelHex = true;
            DisplayColor = new Color();
            RgbColor = new Color();
            SchemeColor = A.SchemeColorValues.Light1;
            Tint = 0;
            Transparency = 0;
        }

        /// <summary>
        /// </summary>
        /// <param name="Color"></param>
        /// <param name="Transparency">Transparency of the color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        internal void SetColor(Color Color, decimal Transparency)
        {
            IsRgbColorModelHex = true;
            RgbColor = Color;
            this.Transparency = Transparency;

            DisplayColor = Color;
        }

        /// <summary>
        /// </summary>
        /// <param name="Color">The theme color used.</param>
        /// <param name="Tint">
        ///     The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color
        ///     and positive tints lighten the theme color.
        /// </param>
        /// <param name="Transparency"></param>
        internal void SetColor(SLThemeColorIndexValues Color, double Tint, decimal Transparency)
        {
            IsRgbColorModelHex = false;
            switch (Color)
            {
                case SLThemeColorIndexValues.Dark1Color:
                    SchemeColor = A.SchemeColorValues.Dark1;
                    break;
                case SLThemeColorIndexValues.Light1Color:
                    SchemeColor = A.SchemeColorValues.Light1;
                    break;
                case SLThemeColorIndexValues.Dark2Color:
                    SchemeColor = A.SchemeColorValues.Dark2;
                    break;
                case SLThemeColorIndexValues.Light2Color:
                    SchemeColor = A.SchemeColorValues.Light2;
                    break;
                case SLThemeColorIndexValues.Accent1Color:
                    SchemeColor = A.SchemeColorValues.Accent1;
                    break;
                case SLThemeColorIndexValues.Accent2Color:
                    SchemeColor = A.SchemeColorValues.Accent2;
                    break;
                case SLThemeColorIndexValues.Accent3Color:
                    SchemeColor = A.SchemeColorValues.Accent3;
                    break;
                case SLThemeColorIndexValues.Accent4Color:
                    SchemeColor = A.SchemeColorValues.Accent4;
                    break;
                case SLThemeColorIndexValues.Accent5Color:
                    SchemeColor = A.SchemeColorValues.Accent5;
                    break;
                case SLThemeColorIndexValues.Accent6Color:
                    SchemeColor = A.SchemeColorValues.Accent6;
                    break;
                case SLThemeColorIndexValues.Hyperlink:
                    SchemeColor = A.SchemeColorValues.Hyperlink;
                    break;
                case SLThemeColorIndexValues.FollowedHyperlinkColor:
                    SchemeColor = A.SchemeColorValues.FollowedHyperlink;
                    break;
            }
            this.Tint = (decimal) Tint;
            this.Transparency = Transparency;

            var index = (int) Color;
            if ((index >= 0) && (index < listThemeColors.Count))
            {
                DisplayColor = System.Drawing.Color.FromArgb(255, listThemeColors[index]);
                if (this.Tint != 0)
                    DisplayColor = SLTool.ToColor(DisplayColor, Tint);
            }
        }

        internal void SetColor(A.SchemeColorValues Color, decimal Tint, decimal Transparency)
        {
            IsRgbColorModelHex = false;

            SchemeColor = Color;

            var iThemeColor = (int) SLThemeColorIndexValues.Dark1Color;
            switch (Color)
            {
                // I don't really know what to assign for Text1, Text2, Background1, Background2
                // PhClr (placeholder colour)
                case A.SchemeColorValues.Dark1:
                case A.SchemeColorValues.Text1:
                    iThemeColor = (int) SLThemeColorIndexValues.Dark1Color;
                    break;
                case A.SchemeColorValues.Light1:
                case A.SchemeColorValues.Background1:
                    iThemeColor = (int) SLThemeColorIndexValues.Light1Color;
                    break;
                case A.SchemeColorValues.Dark2:
                case A.SchemeColorValues.Text2:
                    iThemeColor = (int) SLThemeColorIndexValues.Dark2Color;
                    break;
                case A.SchemeColorValues.Light2:
                case A.SchemeColorValues.Background2:
                    iThemeColor = (int) SLThemeColorIndexValues.Light2Color;
                    break;
                case A.SchemeColorValues.PhColor:
                    iThemeColor = (int) SLThemeColorIndexValues.Accent1Color;
                    break;
                case A.SchemeColorValues.Accent1:
                    iThemeColor = (int) SLThemeColorIndexValues.Accent1Color;
                    break;
                case A.SchemeColorValues.Accent2:
                    iThemeColor = (int) SLThemeColorIndexValues.Accent2Color;
                    break;
                case A.SchemeColorValues.Accent3:
                    iThemeColor = (int) SLThemeColorIndexValues.Accent3Color;
                    break;
                case A.SchemeColorValues.Accent4:
                    iThemeColor = (int) SLThemeColorIndexValues.Accent4Color;
                    break;
                case A.SchemeColorValues.Accent5:
                    iThemeColor = (int) SLThemeColorIndexValues.Accent5Color;
                    break;
                case A.SchemeColorValues.Accent6:
                    iThemeColor = (int) SLThemeColorIndexValues.Accent6Color;
                    break;
                case A.SchemeColorValues.Hyperlink:
                    iThemeColor = (int) SLThemeColorIndexValues.Hyperlink;
                    break;
                case A.SchemeColorValues.FollowedHyperlink:
                    iThemeColor = (int) SLThemeColorIndexValues.FollowedHyperlinkColor;
                    break;
            }
            this.Tint = Tint;
            this.Transparency = Transparency;

            var index = iThemeColor;
            if ((index >= 0) && (index < listThemeColors.Count))
            {
                DisplayColor = System.Drawing.Color.FromArgb(255, listThemeColors[index]);
                if (this.Tint != 0)
                    DisplayColor = SLTool.ToColor(DisplayColor, (double) Tint);
            }
        }

        internal A.RgbColorModelHex ToRgbColorModelHex()
        {
            var rgb = new A.RgbColorModelHex();
            rgb.Val = string.Format("{0}{1}{2}", RgbColor.R.ToString("X2"), RgbColor.G.ToString("X2"),
                RgbColor.B.ToString("X2"));

            var decTint = Tint;

            // we don't have to do anything extra if the tint's zero.
            if (decTint < 0.0m)
            {
                decTint += 1.0m;
                decTint *= 100000m;
                rgb.Append(new A.LuminanceModulation {Val = Convert.ToInt32(decTint)});
            }
            else if (decTint > 0.0m)
            {
                decTint *= 100000m;
                decTint = decimal.Floor(decTint);
                rgb.Append(new A.LuminanceModulation {Val = Convert.ToInt32(100000m - decTint)});
                rgb.Append(new A.LuminanceOffset {Val = Convert.ToInt32(decTint)});
            }

            var iAlpha = SLDrawingTool.CalculateAlpha(Transparency);
            // if >= 100000, then transparency was 0 (or negative),
            // then we don't have to append the Alpha class
            if (iAlpha < 100000)
                rgb.Append(new A.Alpha {Val = iAlpha});

            return rgb;
        }

        internal A.SchemeColor ToSchemeColor()
        {
            var sclr = new A.SchemeColor();
            sclr.Val = SchemeColor;

            var decTint = Tint;

            // we don't have to do anything extra if the tint's zero.
            if (decTint < 0.0m)
            {
                decTint += 1.0m;
                decTint *= 100000m;
                sclr.Append(new A.LuminanceModulation {Val = Convert.ToInt32(decTint)});
            }
            else if (decTint > 0.0m)
            {
                decTint *= 100000m;
                decTint = decimal.Floor(decTint);
                sclr.Append(new A.LuminanceModulation {Val = Convert.ToInt32(100000m - decTint)});
                sclr.Append(new A.LuminanceOffset {Val = Convert.ToInt32(decTint)});
            }

            var iAlpha = SLDrawingTool.CalculateAlpha(Transparency);
            // if >= 100000, then transparency was 0 (or negative),
            // then we don't have to append the Alpha class
            if (iAlpha < 100000)
                sclr.Append(new A.Alpha {Val = iAlpha});

            return sclr;
        }

        internal SLColorTransform Clone()
        {
            var clr = new SLColorTransform(listThemeColors);
            clr.IsRgbColorModelHex = IsRgbColorModelHex;
            clr.DisplayColor = DisplayColor;
            clr.RgbColor = RgbColor;
            clr.SchemeColor = SchemeColor;
            clr.decTint = decTint;
            clr.decTransparency = decTransparency;

            return clr;
        }
    }
}