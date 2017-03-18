using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLightWrapper.Core.misc;
using Color = System.Drawing.Color;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace SpreadsheetLightWrapper.Core.style
{
    /// <summary>
    ///     Encapsulates properties and methods for setting a color. This includes using theme colors. This simulates the
    ///     DocumentFormat.OpenXml.Spreadsheet.Color class.
    /// </summary>
    public class SLColor
    {
        private Color clrDisplay;

        internal double? fTint;
        internal List<Color> listIndexedColors;
        internal List<Color> listThemeColors;

        internal SLColor(List<Color> ThemeColors, List<Color> IndexedColors)
        {
            int i;
            listThemeColors = new List<Color>();
            for (i = 0; i < ThemeColors.Count; ++i)
                listThemeColors.Add(ThemeColors[i]);

            listIndexedColors = new List<Color>();
            for (i = 0; i < IndexedColors.Count; ++i)
                listIndexedColors.Add(IndexedColors[i]);

            SetAllNull();
        }

        /// <summary>
        ///     The color value.
        /// </summary>
        public Color Color
        {
            get { return clrDisplay; }
            set
            {
                SetAllNull();
                clrDisplay = value;
                Rgb = clrDisplay.ToArgb().ToString("X8");
            }
        }

        internal bool? Auto { get; set; }
        internal uint? Indexed { get; set; }
        internal string Rgb { get; set; }
        internal uint? Theme { get; set; }

        internal double? Tint
        {
            get { return fTint; }
            set
            {
                fTint = value;
                if (fTint != null)
                {
                    if (fTint.Value < -1.0) fTint = -1.0;
                    if (fTint.Value > 1.0) fTint = 1.0;
                }
            }
        }

        private void SetAllNull()
        {
            clrDisplay = new Color();
            Auto = null;
            Indexed = null;
            Rgb = null;
            Theme = null;
            Tint = null;
        }

        /// <summary>
        ///     Whether the color value is empty.
        /// </summary>
        /// <returns>True if the color value is empty. False otherwise.</returns>
        public bool IsEmpty()
        {
            return (Auto == null) && (Indexed == null) && (Rgb == null) && (Theme == null);
        }

        /// <summary>
        ///     Set a color using a theme color.
        /// </summary>
        /// <param name="ThemeColorIndex">The theme color to be used.</param>
        public void SetThemeColor(SLThemeColorIndexValues ThemeColorIndex)
        {
            var index = (int) ThemeColorIndex;
            if ((index >= 0) && (index < listThemeColors.Count))
                clrDisplay = listThemeColors[index];
            Auto = null;
            Indexed = null;
            Rgb = null;
            Theme = (uint) ThemeColorIndex;
            Tint = null;
        }

        /// <summary>
        ///     Set a color using a theme color, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="ThemeColorIndex">The theme color to be used.</param>
        /// <param name="Tint">
        ///     The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color
        ///     and positive tints lighten the theme color.
        /// </param>
        public void SetThemeColor(SLThemeColorIndexValues ThemeColorIndex, double Tint)
        {
            var clrRgb = new Color();
            var index = (int) ThemeColorIndex;
            if ((index >= 0) && (index < listThemeColors.Count))
                clrRgb = listThemeColors[index];
            Auto = null;
            Indexed = null;
            Rgb = null;
            Theme = (uint) ThemeColorIndex;
            this.Tint = Tint;
            clrDisplay = SLTool.ToColor(clrRgb, Tint);
        }

        internal BackgroundColor ToBackgroundColor()
        {
            var clr = new BackgroundColor();
            if (Auto != null) clr.Auto = Auto.Value;
            if (Indexed != null) clr.Indexed = Indexed.Value;
            if (Rgb != null) clr.Rgb = new HexBinaryValue(Rgb);
            if (Theme != null) clr.Theme = Theme.Value;
            if ((Tint != null) && (Tint.Value != 0.0)) clr.Tint = Tint.Value;

            return clr;
        }

        internal ForegroundColor ToForegroundColor()
        {
            var clr = new ForegroundColor();
            if (Auto != null) clr.Auto = Auto.Value;
            if (Indexed != null) clr.Indexed = Indexed.Value;
            if (Rgb != null) clr.Rgb = new HexBinaryValue(Rgb);
            if (Theme != null) clr.Theme = Theme.Value;
            if ((Tint != null) && (Tint.Value != 0.0)) clr.Tint = Tint.Value;

            return clr;
        }

        internal TabColor ToTabColor()
        {
            var clr = new TabColor();
            if (Auto != null) clr.Auto = Auto.Value;
            if (Indexed != null) clr.Indexed = Indexed.Value;
            if (Rgb != null) clr.Rgb = new HexBinaryValue(Rgb);
            if (Theme != null) clr.Theme = Theme.Value;
            if ((Tint != null) && (Tint.Value != 0.0)) clr.Tint = Tint.Value;

            return clr;
        }

        internal DocumentFormat.OpenXml.Spreadsheet.Color ToSpreadsheetColor()
        {
            var clr = new DocumentFormat.OpenXml.Spreadsheet.Color();
            if (Auto != null) clr.Auto = Auto.Value;
            if (Indexed != null) clr.Indexed = Indexed.Value;
            if (Rgb != null) clr.Rgb = new HexBinaryValue(Rgb);
            if (Theme != null) clr.Theme = Theme.Value;
            if ((Tint != null) && (Tint.Value != 0.0)) clr.Tint = Tint.Value;

            return clr;
        }

        internal X14.AxisColor ToAxisColor()
        {
            var clr = new X14.AxisColor();
            if (Auto != null) clr.Auto = Auto.Value;
            if (Indexed != null) clr.Indexed = Indexed.Value;
            if (Rgb != null) clr.Rgb = new HexBinaryValue(Rgb);
            if (Theme != null) clr.Theme = Theme.Value;
            if ((Tint != null) && (Tint.Value != 0.0)) clr.Tint = Tint.Value;

            return clr;
        }

        internal X14.BarAxisColor ToBarAxisColor()
        {
            var clr = new X14.BarAxisColor();
            if (Auto != null) clr.Auto = Auto.Value;
            if (Indexed != null) clr.Indexed = Indexed.Value;
            if (Rgb != null) clr.Rgb = new HexBinaryValue(Rgb);
            if (Theme != null) clr.Theme = Theme.Value;
            if ((Tint != null) && (Tint.Value != 0.0)) clr.Tint = Tint.Value;

            return clr;
        }

        internal X14.BorderColor ToBorderColor()
        {
            var clr = new X14.BorderColor();
            if (Auto != null) clr.Auto = Auto.Value;
            if (Indexed != null) clr.Indexed = Indexed.Value;
            if (Rgb != null) clr.Rgb = new HexBinaryValue(Rgb);
            if (Theme != null) clr.Theme = Theme.Value;
            if ((Tint != null) && (Tint.Value != 0.0)) clr.Tint = Tint.Value;

            return clr;
        }

        internal X14.Color ToExcel2010Color()
        {
            var clr = new X14.Color();
            if (Auto != null) clr.Auto = Auto.Value;
            if (Indexed != null) clr.Indexed = Indexed.Value;
            if (Rgb != null) clr.Rgb = new HexBinaryValue(Rgb);
            if (Theme != null) clr.Theme = Theme.Value;
            if ((Tint != null) && (Tint.Value != 0.0)) clr.Tint = Tint.Value;

            return clr;
        }

        internal X14.FillColor ToFillColor()
        {
            var clr = new X14.FillColor();
            if (Auto != null) clr.Auto = Auto.Value;
            if (Indexed != null) clr.Indexed = Indexed.Value;
            if (Rgb != null) clr.Rgb = new HexBinaryValue(Rgb);
            if (Theme != null) clr.Theme = Theme.Value;
            if ((Tint != null) && (Tint.Value != 0.0)) clr.Tint = Tint.Value;

            return clr;
        }

        internal X14.FirstMarkerColor ToFirstMarkerColor()
        {
            var clr = new X14.FirstMarkerColor();
            if (Auto != null) clr.Auto = Auto.Value;
            if (Indexed != null) clr.Indexed = Indexed.Value;
            if (Rgb != null) clr.Rgb = new HexBinaryValue(Rgb);
            if (Theme != null) clr.Theme = Theme.Value;
            if ((Tint != null) && (Tint.Value != 0.0)) clr.Tint = Tint.Value;

            return clr;
        }

        internal X14.HighMarkerColor ToHighMarkerColor()
        {
            var clr = new X14.HighMarkerColor();
            if (Auto != null) clr.Auto = Auto.Value;
            if (Indexed != null) clr.Indexed = Indexed.Value;
            if (Rgb != null) clr.Rgb = new HexBinaryValue(Rgb);
            if (Theme != null) clr.Theme = Theme.Value;
            if ((Tint != null) && (Tint.Value != 0.0)) clr.Tint = Tint.Value;

            return clr;
        }

        internal X14.LastMarkerColor ToLastMarkerColor()
        {
            var clr = new X14.LastMarkerColor();
            if (Auto != null) clr.Auto = Auto.Value;
            if (Indexed != null) clr.Indexed = Indexed.Value;
            if (Rgb != null) clr.Rgb = new HexBinaryValue(Rgb);
            if (Theme != null) clr.Theme = Theme.Value;
            if ((Tint != null) && (Tint.Value != 0.0)) clr.Tint = Tint.Value;

            return clr;
        }

        internal X14.LowMarkerColor ToLowMarkerColor()
        {
            var clr = new X14.LowMarkerColor();
            if (Auto != null) clr.Auto = Auto.Value;
            if (Indexed != null) clr.Indexed = Indexed.Value;
            if (Rgb != null) clr.Rgb = new HexBinaryValue(Rgb);
            if (Theme != null) clr.Theme = Theme.Value;
            if ((Tint != null) && (Tint.Value != 0.0)) clr.Tint = Tint.Value;

            return clr;
        }

        internal X14.MarkersColor ToMarkersColor()
        {
            var clr = new X14.MarkersColor();
            if (Auto != null) clr.Auto = Auto.Value;
            if (Indexed != null) clr.Indexed = Indexed.Value;
            if (Rgb != null) clr.Rgb = new HexBinaryValue(Rgb);
            if (Theme != null) clr.Theme = Theme.Value;
            if ((Tint != null) && (Tint.Value != 0.0)) clr.Tint = Tint.Value;

            return clr;
        }

        internal X14.NegativeBorderColor ToNegativeBorderColor()
        {
            var clr = new X14.NegativeBorderColor();
            if (Auto != null) clr.Auto = Auto.Value;
            if (Indexed != null) clr.Indexed = Indexed.Value;
            if (Rgb != null) clr.Rgb = new HexBinaryValue(Rgb);
            if (Theme != null) clr.Theme = Theme.Value;
            if ((Tint != null) && (Tint.Value != 0.0)) clr.Tint = Tint.Value;

            return clr;
        }

        internal X14.NegativeColor ToNegativeColor()
        {
            var clr = new X14.NegativeColor();
            if (Auto != null) clr.Auto = Auto.Value;
            if (Indexed != null) clr.Indexed = Indexed.Value;
            if (Rgb != null) clr.Rgb = new HexBinaryValue(Rgb);
            if (Theme != null) clr.Theme = Theme.Value;
            if ((Tint != null) && (Tint.Value != 0.0)) clr.Tint = Tint.Value;

            return clr;
        }

        internal X14.NegativeFillColor ToNegativeFillColor()
        {
            var clr = new X14.NegativeFillColor();
            if (Auto != null) clr.Auto = Auto.Value;
            if (Indexed != null) clr.Indexed = Indexed.Value;
            if (Rgb != null) clr.Rgb = new HexBinaryValue(Rgb);
            if (Theme != null) clr.Theme = Theme.Value;
            if ((Tint != null) && (Tint.Value != 0.0)) clr.Tint = Tint.Value;

            return clr;
        }

        internal X14.SeriesColor ToSeriesColor()
        {
            var clr = new X14.SeriesColor();
            if (Auto != null) clr.Auto = Auto.Value;
            if (Indexed != null) clr.Indexed = Indexed.Value;
            if (Rgb != null) clr.Rgb = new HexBinaryValue(Rgb);
            if (Theme != null) clr.Theme = Theme.Value;
            if ((Tint != null) && (Tint.Value != 0.0)) clr.Tint = Tint.Value;

            return clr;
        }

        internal void FromBackgroundColor(BackgroundColor clr)
        {
            SetAllNull();
            if (clr.Auto != null) Auto = clr.Auto.Value;
            if (clr.Indexed != null) Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) Rgb = clr.Rgb.Value;
            if (clr.Theme != null) Theme = clr.Theme.Value;
            if (clr.Tint != null) Tint = clr.Tint.Value;
            SetDisplayColor();
        }

        internal void FromForegroundColor(ForegroundColor clr)
        {
            SetAllNull();
            if (clr.Auto != null) Auto = clr.Auto.Value;
            if (clr.Indexed != null) Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) Rgb = clr.Rgb.Value;
            if (clr.Theme != null) Theme = clr.Theme.Value;
            if (clr.Tint != null) Tint = clr.Tint.Value;
            SetDisplayColor();
        }

        internal void FromTabColor(TabColor clr)
        {
            SetAllNull();
            if (clr.Auto != null) Auto = clr.Auto.Value;
            if (clr.Indexed != null) Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) Rgb = clr.Rgb.Value;
            if (clr.Theme != null) Theme = clr.Theme.Value;
            if (clr.Tint != null) Tint = clr.Tint.Value;
            SetDisplayColor();
        }

        internal void FromSpreadsheetColor(DocumentFormat.OpenXml.Spreadsheet.Color clr)
        {
            SetAllNull();
            if (clr.Auto != null) Auto = clr.Auto.Value;
            if (clr.Indexed != null) Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) Rgb = clr.Rgb.Value;
            if (clr.Theme != null) Theme = clr.Theme.Value;
            if (clr.Tint != null) Tint = clr.Tint.Value;
            SetDisplayColor();
        }

        internal void FromAxisColor(X14.AxisColor clr)
        {
            SetAllNull();
            if (clr.Auto != null) Auto = clr.Auto.Value;
            if (clr.Indexed != null) Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) Rgb = clr.Rgb.Value;
            if (clr.Theme != null) Theme = clr.Theme.Value;
            if (clr.Tint != null) Tint = clr.Tint.Value;
            SetDisplayColor();
        }

        internal void FromBarAxisColor(X14.BarAxisColor clr)
        {
            SetAllNull();
            if (clr.Auto != null) Auto = clr.Auto.Value;
            if (clr.Indexed != null) Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) Rgb = clr.Rgb.Value;
            if (clr.Theme != null) Theme = clr.Theme.Value;
            if (clr.Tint != null) Tint = clr.Tint.Value;
            SetDisplayColor();
        }

        internal void FromBorderColor(X14.BorderColor clr)
        {
            SetAllNull();
            if (clr.Auto != null) Auto = clr.Auto.Value;
            if (clr.Indexed != null) Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) Rgb = clr.Rgb.Value;
            if (clr.Theme != null) Theme = clr.Theme.Value;
            if (clr.Tint != null) Tint = clr.Tint.Value;
            SetDisplayColor();
        }

        internal void FromExcel2010Color(X14.Color clr)
        {
            SetAllNull();
            if (clr.Auto != null) Auto = clr.Auto.Value;
            if (clr.Indexed != null) Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) Rgb = clr.Rgb.Value;
            if (clr.Theme != null) Theme = clr.Theme.Value;
            if (clr.Tint != null) Tint = clr.Tint.Value;
            SetDisplayColor();
        }

        internal void FromFillColor(X14.FillColor clr)
        {
            SetAllNull();
            if (clr.Auto != null) Auto = clr.Auto.Value;
            if (clr.Indexed != null) Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) Rgb = clr.Rgb.Value;
            if (clr.Theme != null) Theme = clr.Theme.Value;
            if (clr.Tint != null) Tint = clr.Tint.Value;
            SetDisplayColor();
        }

        internal void FromFirstMarkerColor(X14.FirstMarkerColor clr)
        {
            SetAllNull();
            if (clr.Auto != null) Auto = clr.Auto.Value;
            if (clr.Indexed != null) Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) Rgb = clr.Rgb.Value;
            if (clr.Theme != null) Theme = clr.Theme.Value;
            if (clr.Tint != null) Tint = clr.Tint.Value;
            SetDisplayColor();
        }

        internal void FromHighMarkerColor(X14.HighMarkerColor clr)
        {
            SetAllNull();
            if (clr.Auto != null) Auto = clr.Auto.Value;
            if (clr.Indexed != null) Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) Rgb = clr.Rgb.Value;
            if (clr.Theme != null) Theme = clr.Theme.Value;
            if (clr.Tint != null) Tint = clr.Tint.Value;
            SetDisplayColor();
        }

        internal void FromLastMarkerColor(X14.LastMarkerColor clr)
        {
            SetAllNull();
            if (clr.Auto != null) Auto = clr.Auto.Value;
            if (clr.Indexed != null) Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) Rgb = clr.Rgb.Value;
            if (clr.Theme != null) Theme = clr.Theme.Value;
            if (clr.Tint != null) Tint = clr.Tint.Value;
            SetDisplayColor();
        }

        internal void FromLowMarkerColor(X14.LowMarkerColor clr)
        {
            SetAllNull();
            if (clr.Auto != null) Auto = clr.Auto.Value;
            if (clr.Indexed != null) Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) Rgb = clr.Rgb.Value;
            if (clr.Theme != null) Theme = clr.Theme.Value;
            if (clr.Tint != null) Tint = clr.Tint.Value;
            SetDisplayColor();
        }

        internal void FromMarkersColor(X14.MarkersColor clr)
        {
            SetAllNull();
            if (clr.Auto != null) Auto = clr.Auto.Value;
            if (clr.Indexed != null) Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) Rgb = clr.Rgb.Value;
            if (clr.Theme != null) Theme = clr.Theme.Value;
            if (clr.Tint != null) Tint = clr.Tint.Value;
            SetDisplayColor();
        }

        internal void FromNegativeBorderColor(X14.NegativeBorderColor clr)
        {
            SetAllNull();
            if (clr.Auto != null) Auto = clr.Auto.Value;
            if (clr.Indexed != null) Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) Rgb = clr.Rgb.Value;
            if (clr.Theme != null) Theme = clr.Theme.Value;
            if (clr.Tint != null) Tint = clr.Tint.Value;
            SetDisplayColor();
        }

        internal void FromNegativeColor(X14.NegativeColor clr)
        {
            SetAllNull();
            if (clr.Auto != null) Auto = clr.Auto.Value;
            if (clr.Indexed != null) Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) Rgb = clr.Rgb.Value;
            if (clr.Theme != null) Theme = clr.Theme.Value;
            if (clr.Tint != null) Tint = clr.Tint.Value;
            SetDisplayColor();
        }

        internal void FromNegativeFillColor(X14.NegativeFillColor clr)
        {
            SetAllNull();
            if (clr.Auto != null) Auto = clr.Auto.Value;
            if (clr.Indexed != null) Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) Rgb = clr.Rgb.Value;
            if (clr.Theme != null) Theme = clr.Theme.Value;
            if (clr.Tint != null) Tint = clr.Tint.Value;
            SetDisplayColor();
        }

        internal void FromSeriesColor(X14.SeriesColor clr)
        {
            SetAllNull();
            if (clr.Auto != null) Auto = clr.Auto.Value;
            if (clr.Indexed != null) Indexed = clr.Indexed.Value;
            if (clr.Rgb != null) Rgb = clr.Rgb.Value;
            if (clr.Theme != null) Theme = clr.Theme.Value;
            if (clr.Tint != null) Tint = clr.Tint.Value;
            SetDisplayColor();
        }

        private void SetDisplayColor()
        {
            clrDisplay = Color.FromArgb(255, 0, 0, 0);

            var index = 0;
            if (Theme != null)
            {
                index = (int) Theme.Value;
                if ((index >= 0) && (index < listThemeColors.Count))
                {
                    clrDisplay = Color.FromArgb(255, listThemeColors[index]);
                    if (Tint != null)
                        clrDisplay = SLTool.ToColor(clrDisplay, Tint.Value);
                }
            }
            else if (Rgb != null)
            {
                clrDisplay = SLTool.ToColor(Rgb);
            }
            else if (Indexed != null)
            {
                index = (int) Indexed.Value;
                if ((index >= 0) && (index < listIndexedColors.Count))
                    clrDisplay = Color.FromArgb(255, listIndexedColors[index]);
            }
        }

        internal void FromHash(string Hash)
        {
            SetAllNull();

            var sa = Hash.Split(new[] {SLConstants.XmlColorAttributeSeparator}, StringSplitOptions.None);
            if (sa.Length >= 5)
            {
                if (!sa[0].Equals("null")) Auto = bool.Parse(sa[0]);

                if (!sa[1].Equals("null")) Indexed = uint.Parse(sa[1]);

                if (!sa[2].Equals("null")) Rgb = sa[2];

                if (!sa[3].Equals("null")) Theme = uint.Parse(sa[3]);

                if (!sa[4].Equals("null")) Tint = double.Parse(sa[4]);
            }

            SetDisplayColor();
        }

        internal string ToHash()
        {
            var sb = new StringBuilder();

            if (Auto != null) sb.AppendFormat("{0}{1}", Auto.Value, SLConstants.XmlColorAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlColorAttributeSeparator);

            if (Indexed != null) sb.AppendFormat("{0}{1}", Indexed.Value, SLConstants.XmlColorAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlColorAttributeSeparator);

            if (Rgb != null) sb.AppendFormat("{0}{1}", Rgb, SLConstants.XmlColorAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlColorAttributeSeparator);

            if (Theme != null) sb.AppendFormat("{0}{1}", Theme.Value, SLConstants.XmlColorAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlColorAttributeSeparator);

            if (Tint != null) sb.AppendFormat("{0}{1}", Tint.Value, SLConstants.XmlColorAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlColorAttributeSeparator);

            return sb.ToString();
        }

        internal SLColor Clone()
        {
            var clr = new SLColor(listThemeColors, listIndexedColors);
            clr.clrDisplay = clrDisplay;
            clr.Auto = Auto;
            clr.Indexed = Indexed;
            clr.Rgb = Rgb;
            clr.Theme = Theme;
            clr.Tint = Tint;

            return clr;
        }
    }
}