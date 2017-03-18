using System;
using System.Collections.Generic;
using System.Drawing;
using Ups.Toolkit.SpreadsheetLight.Core.misc;
using A = DocumentFormat.OpenXml.Drawing;

namespace Ups.Toolkit.SpreadsheetLight.Core.Drawing
{
    /// <summary>
    ///     Encapsulates properties and methods for setting line or border settings.
    ///     This simulates the DocumentFormat.OpenXml.Drawing.LinePropertiesType class.
    /// </summary>
    public class SLLinePropertiesType
    {
        private bool bUseGradientLine;

        private bool bUseNoLine;

        private bool bUseSolidLine;
        private decimal decWidth;

        internal bool HasCapType;

        internal bool HasCompoundLineType;

        internal bool HasDashType;

        internal bool HasJoinType;

        internal bool HasWidth;
        internal List<Color> listThemeColors;
        private A.LineCapValues vCapType;
        private A.CompoundLineValues vCompoundLineType;
        private A.PresetLineDashValues vDashType;
        private SLLineJoinValues vJoinType;

        internal SLLinePropertiesType(List<Color> ThemeColors)
        {
            int i;
            listThemeColors = new List<Color>();
            for (i = 0; i < ThemeColors.Count; ++i)
                listThemeColors.Add(ThemeColors[i]);

            SetAllNull();
        }

        internal bool HasLine
        {
            get
            {
                return UseNoLine || UseSolidLine || UseGradientLine || HasWidth || HasCapType || HasCompoundLineType ||
                       HasDashType || HasJoinType;
            }
        }

        internal bool UseNoLine
        {
            get { return bUseNoLine; }
            set
            {
                bUseNoLine = value;
                if (value)
                {
                    bUseNoLine = true;
                    bUseSolidLine = false;
                    bUseGradientLine = false;
                }
            }
        }

        internal bool UseSolidLine
        {
            get { return bUseSolidLine; }
            set
            {
                bUseSolidLine = value;
                if (value)
                {
                    bUseNoLine = false;
                    bUseSolidLine = true;
                    bUseGradientLine = false;
                }
            }
        }

        internal SLColorTransform SolidColor { get; set; }

        internal bool UseGradientLine
        {
            get { return bUseGradientLine; }
            set
            {
                bUseGradientLine = value;
                if (value)
                {
                    bUseNoLine = false;
                    bUseSolidLine = false;
                    bUseGradientLine = true;
                }
            }
        }

        internal SLGradientFill GradientColor { get; set; }

        /// <summary>
        ///     The dash type.
        /// </summary>
        public A.PresetLineDashValues DashType
        {
            get { return vDashType; }
            set
            {
                HasDashType = true;
                vDashType = value;
            }
        }

        /// <summary>
        ///     The join type.
        /// </summary>
        public SLLineJoinValues JoinType
        {
            get { return vJoinType; }
            set
            {
                HasJoinType = true;
                vJoinType = value;
            }
        }

        internal A.LineEndValues? HeadEndType { get; set; }
        internal SLLineSizeValues HeadEndSize { get; set; }
        internal A.LineEndValues? TailEndType { get; set; }
        internal SLLineSizeValues TailEndSize { get; set; }

        /// <summary>
        ///     Width between 0 pt and 1584 pt. Accurate to 1/12700 of a point.
        /// </summary>
        public decimal Width
        {
            get { return decWidth; }
            set
            {
                HasWidth = true;
                decWidth = value;
                if (decWidth < 0m) decWidth = 0m;
                if (decWidth > 1584m) decWidth = 1584m;
            }
        }

        /// <summary>
        ///     The cap type.
        /// </summary>
        public A.LineCapValues CapType
        {
            get { return vCapType; }
            set
            {
                HasCapType = true;
                vCapType = value;
            }
        }

        /// <summary>
        ///     The compound type.
        /// </summary>
        public A.CompoundLineValues CompoundLineType
        {
            get { return vCompoundLineType; }
            set
            {
                HasCompoundLineType = true;
                vCompoundLineType = value;
            }
        }

        /// <summary>
        ///     The alignment.
        /// </summary>
        public A.PenAlignmentValues? Alignment { get; set; }

        private void SetAllNull()
        {
            bUseNoLine = false;
            bUseSolidLine = false;
            SolidColor = new SLColorTransform(listThemeColors);
            bUseGradientLine = false;
            GradientColor = new SLGradientFill(listThemeColors);

            decWidth = 0m;
            HasWidth = false;
            vCompoundLineType = A.CompoundLineValues.Single;
            HasCompoundLineType = false;
            vDashType = A.PresetLineDashValues.Solid;
            HasDashType = false;
            vCapType = A.LineCapValues.Square;
            HasCapType = false;
            vJoinType = SLLineJoinValues.Round;
            HasJoinType = false;

            HeadEndType = null;
            HeadEndSize = SLLineSizeValues.Size1;
            TailEndType = null;
            TailEndSize = SLLineSizeValues.Size1;

            Alignment = null;
        }

        /// <summary>
        ///     Set color to be automatic.
        /// </summary>
        public void SetAutomaticColor()
        {
            bUseNoLine = false;
            bUseSolidLine = false;
            bUseGradientLine = false;
        }

        /// <summary>
        ///     Set no line.
        /// </summary>
        public void SetNoLine()
        {
            UseNoLine = true;
        }

        /// <summary>
        ///     Set a solid line given a color for the line and the transparency of the color.
        /// </summary>
        /// <param name="Color">The color to be used.</param>
        /// <param name="Transparency">Transparency of the color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        public void SetSolidLine(Color Color, decimal Transparency)
        {
            UseSolidLine = true;
            SolidColor.SetColor(Color, Transparency);
        }

        /// <summary>
        ///     Set a solid line given a color for the line and the transparency of the color.
        /// </summary>
        /// <param name="Color">The theme color to be used.</param>
        /// <param name="Tint">
        ///     The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color
        ///     and positive tints lighten the theme color.
        /// </param>
        /// <param name="Transparency">Transparency of the color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        public void SetSolidLine(SLThemeColorIndexValues Color, double Tint, decimal Transparency)
        {
            UseSolidLine = true;
            SolidColor.SetColor(Color, Tint, Transparency);
        }

        internal void SetSolidLine(A.SchemeColorValues Color, decimal Tint, decimal Transparency)
        {
            UseSolidLine = true;
            SolidColor.SetColor(Color, Tint, Transparency);
        }

        /// <summary>
        ///     Set a linear gradient given a preset setting.
        /// </summary>
        /// <param name="Preset">The preset to be used.</param>
        /// <param name="Angle">
        ///     The interpolation angle ranging from 0 degrees to 359.9 degrees. 0 degrees mean from left to right,
        ///     90 degrees mean from top to bottom, 180 degrees mean from right to left and 270 degrees mean from bottom to top.
        ///     Accurate to 1/60000 of a degree.
        /// </param>
        public void SetLinearGradient(SLGradientPresetValues Preset, decimal Angle)
        {
            UseGradientLine = true;
            GradientColor.SetLinearGradient(Preset, Angle);
        }

        /// <summary>
        ///     Set a radial gradient given a preset setting.
        /// </summary>
        /// <param name="Preset">The preset to be used.</param>
        /// <param name="Direction">The radial gradient direction.</param>
        public void SetRadialGradient(SLGradientPresetValues Preset, SLGradientDirectionValues Direction)
        {
            UseGradientLine = true;
            GradientColor.SetRadialGradient(Preset, Direction);
        }

        /// <summary>
        ///     Set a rectangular gradient given a preset setting.
        /// </summary>
        /// <param name="Preset">The preset to be used.</param>
        /// <param name="Direction">The rectangular gradient direction.</param>
        public void SetRectangularGradient(SLGradientPresetValues Preset, SLGradientDirectionValues Direction)
        {
            UseGradientLine = true;
            GradientColor.SetRectangularGradient(Preset, Direction);
        }

        /// <summary>
        ///     Set a path gradient given a preset setting.
        /// </summary>
        /// <param name="Preset">The preset to be used.</param>
        public void SetPathGradient(SLGradientPresetValues Preset)
        {
            UseGradientLine = true;
            GradientColor.SetPathGradient(Preset);
        }

        /// <summary>
        ///     Append a gradient stop given a color, the color's transparency and the position of gradient stop.
        /// </summary>
        /// <param name="Color">The color to be used.</param>
        /// <param name="Transparency">Transparency of the color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="Position">The position in percentage ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        public void AppendGradientStop(Color Color, decimal Transparency, decimal Position)
        {
            GradientColor.AppendGradientStop(Color, Transparency, Position);
        }

        /// <summary>
        ///     Append a gradient stop given a color, the color's transparency and the position of gradient stop.
        /// </summary>
        /// <param name="Color">The theme color to be used.</param>
        /// <param name="Tint">
        ///     The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color
        ///     and positive tints lighten the theme color.
        /// </param>
        /// <param name="Transparency">Transparency of the color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        /// <param name="Position">The position in percentage ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        public void AppendGradientStop(SLThemeColorIndexValues Color, double Tint, decimal Transparency,
            decimal Position)
        {
            GradientColor.AppendGradientStop(Color, Tint, Transparency, Position);
        }

        /// <summary>
        ///     Clear all gradient stops.
        /// </summary>
        public void ClearGradientStops()
        {
            GradientColor.ClearGradientStops();
        }

        /// <summary>
        ///     Set line arrow head settings. This only makes sense for lines and not border lines.
        /// </summary>
        /// <param name="HeadType">The arrow head type.</param>
        /// <param name="HeadSize">The arrow head size.</param>
        public void SetArrowHead(A.LineEndValues HeadType, SLLineSizeValues HeadSize)
        {
            HeadEndType = HeadType;
            HeadEndSize = HeadSize;
        }

        /// <summary>
        ///     Set line arrow tail settings. This only makes sense for lines and not border lines.
        /// </summary>
        /// <param name="TailType">The arrow tail type.</param>
        /// <param name="TailSize">The arrow tail size.</param>
        public void SetArrowTail(A.LineEndValues TailType, SLLineSizeValues TailSize)
        {
            TailEndType = TailType;
            TailEndSize = TailSize;
        }

        internal A.Outline ToOutline()
        {
            var ol = new A.Outline();
            if (UseNoLine) ol.Append(new A.NoFill());
            if (UseSolidLine)
                if (SolidColor.IsRgbColorModelHex)
                    ol.Append(new A.SolidFill {RgbColorModelHex = SolidColor.ToRgbColorModelHex()});
                else
                    ol.Append(new A.SolidFill {SchemeColor = SolidColor.ToSchemeColor()});
            if (UseGradientLine)
                ol.Append(GradientColor.ToGradientFill());

            if (HasDashType) ol.Append(new A.PresetDash {Val = DashType});

            if (HasJoinType)
                switch (JoinType)
                {
                    case SLLineJoinValues.Round:
                        ol.Append(new A.Round());
                        break;
                    case SLLineJoinValues.Bevel:
                        ol.Append(new A.Bevel());
                        break;
                    case SLLineJoinValues.Miter:
                        // 800000 was the default Excel gave
                        ol.Append(new A.Miter {Limit = 800000});
                        break;
                }

            if (HeadEndType != null) ol.Append(GetHeadEnd());
            if (TailEndType != null) ol.Append(GetTailEnd());

            if (HasWidth) ol.Width = Convert.ToInt32(Width*SLConstants.PointToEMU);
            if (HasCapType) ol.CapType = CapType;
            if (HasCompoundLineType) ol.CompoundLineType = CompoundLineType;
            if (Alignment != null) ol.Alignment = Alignment.Value;

            return ol;
        }

        private A.HeadEnd GetHeadEnd()
        {
            var he = new A.HeadEnd {Type = HeadEndType.Value};
            switch (HeadEndSize)
            {
                case SLLineSizeValues.Size1:
                    he.Width = A.LineEndWidthValues.Small;
                    he.Length = A.LineEndLengthValues.Small;
                    break;
                case SLLineSizeValues.Size2:
                    he.Width = A.LineEndWidthValues.Small;
                    he.Length = A.LineEndLengthValues.Medium;
                    break;
                case SLLineSizeValues.Size3:
                    he.Width = A.LineEndWidthValues.Small;
                    he.Length = A.LineEndLengthValues.Large;
                    break;
                case SLLineSizeValues.Size4:
                    he.Width = A.LineEndWidthValues.Medium;
                    he.Length = A.LineEndLengthValues.Small;
                    break;
                case SLLineSizeValues.Size5:
                    he.Width = A.LineEndWidthValues.Medium;
                    he.Length = A.LineEndLengthValues.Medium;
                    break;
                case SLLineSizeValues.Size6:
                    he.Width = A.LineEndWidthValues.Medium;
                    he.Length = A.LineEndLengthValues.Large;
                    break;
                case SLLineSizeValues.Size7:
                    he.Width = A.LineEndWidthValues.Large;
                    he.Length = A.LineEndLengthValues.Small;
                    break;
                case SLLineSizeValues.Size8:
                    he.Width = A.LineEndWidthValues.Large;
                    he.Length = A.LineEndLengthValues.Medium;
                    break;
                case SLLineSizeValues.Size9:
                    he.Width = A.LineEndWidthValues.Large;
                    he.Length = A.LineEndLengthValues.Large;
                    break;
            }

            return he;
        }

        private A.TailEnd GetTailEnd()
        {
            var te = new A.TailEnd {Type = TailEndType.Value};
            switch (TailEndSize)
            {
                case SLLineSizeValues.Size1:
                    te.Width = A.LineEndWidthValues.Small;
                    te.Length = A.LineEndLengthValues.Small;
                    break;
                case SLLineSizeValues.Size2:
                    te.Width = A.LineEndWidthValues.Small;
                    te.Length = A.LineEndLengthValues.Medium;
                    break;
                case SLLineSizeValues.Size3:
                    te.Width = A.LineEndWidthValues.Small;
                    te.Length = A.LineEndLengthValues.Large;
                    break;
                case SLLineSizeValues.Size4:
                    te.Width = A.LineEndWidthValues.Medium;
                    te.Length = A.LineEndLengthValues.Small;
                    break;
                case SLLineSizeValues.Size5:
                    te.Width = A.LineEndWidthValues.Medium;
                    te.Length = A.LineEndLengthValues.Medium;
                    break;
                case SLLineSizeValues.Size6:
                    te.Width = A.LineEndWidthValues.Medium;
                    te.Length = A.LineEndLengthValues.Large;
                    break;
                case SLLineSizeValues.Size7:
                    te.Width = A.LineEndWidthValues.Large;
                    te.Length = A.LineEndLengthValues.Small;
                    break;
                case SLLineSizeValues.Size8:
                    te.Width = A.LineEndWidthValues.Large;
                    te.Length = A.LineEndLengthValues.Medium;
                    break;
                case SLLineSizeValues.Size9:
                    te.Width = A.LineEndWidthValues.Large;
                    te.Length = A.LineEndLengthValues.Large;
                    break;
            }

            return te;
        }

        internal SLLinePropertiesType Clone()
        {
            var lpt = new SLLinePropertiesType(listThemeColors);
            lpt.bUseNoLine = bUseNoLine;
            lpt.bUseSolidLine = bUseSolidLine;
            lpt.SolidColor = SolidColor.Clone();
            lpt.bUseGradientLine = bUseGradientLine;
            lpt.GradientColor = GradientColor.Clone();
            lpt.vDashType = vDashType;
            lpt.HasDashType = HasDashType;
            lpt.vJoinType = vJoinType;
            lpt.HasJoinType = HasJoinType;
            lpt.HeadEndType = HeadEndType;
            lpt.HeadEndSize = HeadEndSize;
            lpt.TailEndType = TailEndType;
            lpt.TailEndSize = TailEndSize;
            lpt.decWidth = decWidth;
            lpt.HasWidth = HasWidth;
            lpt.vCapType = vCapType;
            lpt.HasCapType = HasCapType;
            lpt.vCompoundLineType = vCompoundLineType;
            lpt.HasCompoundLineType = HasCompoundLineType;
            lpt.Alignment = Alignment;

            return lpt;
        }
    }
}