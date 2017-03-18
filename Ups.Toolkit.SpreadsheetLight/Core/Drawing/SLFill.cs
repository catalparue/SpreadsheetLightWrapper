using System.Collections.Generic;
using System.Drawing;
using DocumentFormat.OpenXml;
using Ups.Toolkit.SpreadsheetLight.Core.misc;
using A = DocumentFormat.OpenXml.Drawing;

namespace Ups.Toolkit.SpreadsheetLight.Core.Drawing
{
    internal enum SLFillType
    {
        Automatic = 0,
        NoFill,
        SolidFill,
        GradientFill,
        BlipFill,
        PatternFill
    }

    /// <summary>
    ///     Encapsulates properties and methods for specifying fill effects.
    ///     This simulates the DocumentFormat.OpenXml.Drawing.Fill class.
    /// </summary>
    public class SLFill
    {
        private decimal decBlipBottomOffset;
        private decimal decBlipLeftOffset;
        private decimal decBlipOffsetX;
        private decimal decBlipOffsetY;
        private decimal decBlipRightOffset;
        private decimal decBlipScaleX;
        private decimal decBlipScaleY;
        private decimal decBlipTopOffset;
        private decimal decBlipTransparency;
        internal List<Color> listThemeColors;

        internal SLFillType Type;

        internal SLFill(List<Color> ThemeColors)
        {
            int i;
            listThemeColors = new List<Color>();
            for (i = 0; i < ThemeColors.Count; ++i)
                listThemeColors.Add(ThemeColors[i]);

            SetAllNull();
        }

        internal bool HasFill
        {
            get { return Type != SLFillType.Automatic ? true : false; }
        }

        internal SLColorTransform SolidColor { get; set; }

        internal SLGradientFill GradientColor { get; set; }

        internal string BlipFileName { get; set; }
        internal string BlipRelationshipID { get; set; }
        internal bool BlipTile { get; set; }

        internal decimal BlipLeftOffset
        {
            get { return decBlipLeftOffset; }
            set
            {
                decBlipLeftOffset = value;
                if (decBlipLeftOffset < -100m) decBlipLeftOffset = -100m;
                if (decBlipLeftOffset > 100m) decBlipLeftOffset = 100m;
            }
        }

        internal decimal BlipRightOffset
        {
            get { return decBlipRightOffset; }
            set
            {
                decBlipRightOffset = value;
                if (decBlipRightOffset < -100m) decBlipRightOffset = -100m;
                if (decBlipRightOffset > 100m) decBlipRightOffset = 100m;
            }
        }

        internal decimal BlipTopOffset
        {
            get { return decBlipTopOffset; }
            set
            {
                decBlipTopOffset = value;
                if (decBlipTopOffset < -100m) decBlipTopOffset = -100m;
                if (decBlipTopOffset > 100m) decBlipTopOffset = 100m;
            }
        }

        internal decimal BlipBottomOffset
        {
            get { return decBlipBottomOffset; }
            set
            {
                decBlipBottomOffset = value;
                if (decBlipBottomOffset < -100m) decBlipBottomOffset = -100m;
                if (decBlipBottomOffset > 100m) decBlipBottomOffset = 100m;
            }
        }

        internal decimal BlipOffsetX
        {
            get { return decBlipOffsetX; }
            set
            {
                decBlipOffsetX = value;
                if (decBlipOffsetX < -1584m) decBlipOffsetX = -1584m;
                if (decBlipOffsetX > 1584m) decBlipOffsetX = 1584m;
            }
        }

        internal decimal BlipOffsetY
        {
            get { return decBlipOffsetY; }
            set
            {
                decBlipOffsetY = value;
                if (decBlipOffsetY < -1584m) decBlipOffsetY = -1584m;
                if (decBlipOffsetY > 1584m) decBlipOffsetY = 1584m;
            }
        }

        internal decimal BlipScaleX
        {
            get { return decBlipScaleX; }
            set
            {
                decBlipScaleX = value;
                if (decBlipScaleX < 0m) decBlipScaleX = 0m;
                if (decBlipScaleX > 100m) decBlipScaleX = 100m;
            }
        }

        internal decimal BlipScaleY
        {
            get { return decBlipScaleY; }
            set
            {
                decBlipScaleY = value;
                if (decBlipScaleY < 0m) decBlipScaleY = 0m;
                if (decBlipScaleY > 100m) decBlipScaleY = 100m;
            }
        }

        internal A.RectangleAlignmentValues BlipAlignment { get; set; }
        internal A.TileFlipValues BlipMirrorType { get; set; }

        internal decimal BlipTransparency
        {
            get { return decBlipTransparency; }
            set
            {
                decBlipTransparency = value;
                if (decBlipTransparency < 0m) decBlipTransparency = 0m;
                if (decBlipTransparency > 100m) decBlipTransparency = 100m;
            }
        }

        internal uint? BlipDpi { get; set; }
        internal bool? BlipRotateWithShape { get; set; }

        internal A.PresetPatternValues PatternPreset { get; set; }
        internal SLColorTransform PatternForegroundColor { get; set; }
        internal SLColorTransform PatternBackgroundColor { get; set; }

        private void SetAllNull()
        {
            Type = SLFillType.Automatic;
            SolidColor = new SLColorTransform(listThemeColors);
            GradientColor = new SLGradientFill(listThemeColors);
            BlipFileName = string.Empty;
            BlipRelationshipID = string.Empty;
            BlipTile = true;
            BlipLeftOffset = 0;
            BlipRightOffset = 0;
            BlipTopOffset = 0;
            BlipBottomOffset = 0;
            BlipOffsetX = 0;
            BlipOffsetY = 0;
            BlipScaleX = 100;
            BlipScaleY = 100;
            BlipAlignment = A.RectangleAlignmentValues.TopLeft;
            BlipMirrorType = A.TileFlipValues.None;
            BlipTransparency = 0;
            BlipDpi = null;
            BlipRotateWithShape = null;
            PatternForegroundColor = new SLColorTransform(listThemeColors);
            PatternBackgroundColor = new SLColorTransform(listThemeColors);
        }

        /// <summary>
        ///     Set the fill to automatic.
        /// </summary>
        public void SetAutomaticFill()
        {
            Type = SLFillType.Automatic;
        }

        /// <summary>
        ///     Set no fill.
        /// </summary>
        public void SetNoFill()
        {
            Type = SLFillType.NoFill;
        }

        /// <summary>
        ///     Set a solid fill.
        /// </summary>
        /// <param name="FillColor">The color used.</param>
        /// <param name="Transparency">Transparency of the color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        public void SetSolidFill(Color FillColor, decimal Transparency)
        {
            Type = SLFillType.SolidFill;
            SolidColor.SetColor(FillColor, Transparency);
        }

        /// <summary>
        ///     Set a solid fill.
        /// </summary>
        /// <param name="FillColor">The theme color used.</param>
        /// <param name="Tint">
        ///     The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color
        ///     and positive tints lighten the theme color.
        /// </param>
        /// <param name="Transparency">Transparency of the color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        public void SetSolidFill(SLThemeColorIndexValues FillColor, double Tint, decimal Transparency)
        {
            Type = SLFillType.SolidFill;
            SolidColor.SetColor(FillColor, Tint, Transparency);
        }

        internal void SetSolidFill(A.SchemeColorValues FillColor, decimal Tint, decimal Transparency)
        {
            Type = SLFillType.SolidFill;
            SolidColor.SetColor(FillColor, Tint, Transparency);
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
            Type = SLFillType.GradientFill;
            GradientColor.SetLinearGradient(Preset, Angle);
        }

        /// <summary>
        ///     Set a radial gradient given a preset setting.
        /// </summary>
        /// <param name="Preset">The preset to be used.</param>
        /// <param name="Direction">The radial gradient direction.</param>
        public void SetRadialGradient(SLGradientPresetValues Preset, SLGradientDirectionValues Direction)
        {
            Type = SLFillType.GradientFill;
            GradientColor.SetRadialGradient(Preset, Direction);
        }

        /// <summary>
        ///     Set a rectangular gradient given a preset setting.
        /// </summary>
        /// <param name="Preset">The preset to be used.</param>
        /// <param name="Direction">The rectangular gradient direction.</param>
        public void SetRectangularGradient(SLGradientPresetValues Preset, SLGradientDirectionValues Direction)
        {
            Type = SLFillType.GradientFill;
            GradientColor.SetRectangularGradient(Preset, Direction);
        }

        /// <summary>
        ///     Set a path gradient given a preset setting.
        /// </summary>
        /// <param name="Preset">The preset to be used.</param>
        public void SetPathGradient(SLGradientPresetValues Preset)
        {
            Type = SLFillType.GradientFill;
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
        ///     Set a picture fill. This stretches the picture.
        /// </summary>
        /// <param name="PictureFileName">The file name of the image/picture used.</param>
        /// <param name="LeftOffset">
        ///     The left offset in percentage. A suggested range is -100% to 100%. Accurate to 1/1000 of a
        ///     percent.
        /// </param>
        /// <param name="RightOffset">
        ///     The right offset in percentage. A suggested range is -100% to 100%. Accurate to 1/1000 of a
        ///     percent.
        /// </param>
        /// <param name="TopOffset">
        ///     The top offset in percentage. A suggested range is -100% to 100%. Accurate to 1/1000 of a
        ///     percent.
        /// </param>
        /// <param name="BottomOffset">
        ///     The bottom offset in percentage. A suggested range is -100% to 100%. Accurate to 1/1000 of a
        ///     percent.
        /// </param>
        /// <param name="Transparency">Transparency of the picture ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        public void SetPictureFill(string PictureFileName, decimal LeftOffset, decimal RightOffset, decimal TopOffset,
            decimal BottomOffset, decimal Transparency)
        {
            Type = SLFillType.BlipFill;
            BlipTile = false;
            BlipFileName = PictureFileName;
            BlipLeftOffset = LeftOffset;
            BlipRightOffset = RightOffset;
            BlipTopOffset = TopOffset;
            BlipBottomOffset = BottomOffset;
            BlipTransparency = Transparency;
        }

        /// <summary>
        ///     Set a picture fill. This tiles the picture.
        /// </summary>
        /// <param name="PictureFileName">The file name of the image/picture used.</param>
        /// <param name="OffsetX">
        ///     Horizontal offset ranging from -2147483648 pt to 2147483647 pt. However a suggested range is
        ///     -1585pt to 1584pt. Accurate to 1/12700 of a point.
        /// </param>
        /// <param name="OffsetY">
        ///     Vertical offset ranging from -2147483648 pt to 2147483647 pt. However a suggested range is
        ///     -1585pt to 1584pt. Accurate to 1/12700 of a point.
        /// </param>
        /// <param name="ScaleX">Horizontal scale in percentage. A suggested range is 0% to 100%.</param>
        /// <param name="ScaleY">Vertical scale in percentage. A suggested range is 0% to 100%.</param>
        /// <param name="Alignment">Picture alignment.</param>
        /// <param name="MirrorType">Picture mirror type.</param>
        /// <param name="Transparency">Transparency of the picture ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        public void SetPictureFill(string PictureFileName, decimal OffsetX, decimal OffsetY, decimal ScaleX,
            decimal ScaleY, A.RectangleAlignmentValues Alignment, A.TileFlipValues MirrorType, decimal Transparency)
        {
            Type = SLFillType.BlipFill;
            BlipTile = true;
            BlipFileName = PictureFileName;
            BlipOffsetX = OffsetX;
            BlipOffsetY = OffsetY;
            BlipScaleX = ScaleX;
            BlipScaleY = ScaleY;
            BlipAlignment = Alignment;
            BlipMirrorType = MirrorType;
            BlipTransparency = Transparency;
        }

        /// <summary>
        ///     Set a pattern fill with a preset pattern, foreground color and background color.
        /// </summary>
        /// <param name="PresetPattern">A preset fill pattern.</param>
        /// <param name="ForegroundColor">The color to be used for the foreground.</param>
        /// <param name="BackgroundColor">The color to be used for the background.</param>
        public void SetPatternFill(A.PresetPatternValues PresetPattern, Color ForegroundColor, Color BackgroundColor)
        {
            Type = SLFillType.PatternFill;
            PatternPreset = PresetPattern;
            PatternForegroundColor.SetColor(ForegroundColor, 0);
            PatternBackgroundColor.SetColor(BackgroundColor, 0);
        }

        /// <summary>
        ///     Set a pattern fill with a preset pattern, foreground color and background color.
        /// </summary>
        /// <param name="PresetPattern">A preset fill pattern.</param>
        /// <param name="ForegroundColor">The color to be used for the foreground.</param>
        /// <param name="BackgroundColorTheme">The theme color to be used for the background.</param>
        public void SetPatternFill(A.PresetPatternValues PresetPattern, Color ForegroundColor,
            SLThemeColorIndexValues BackgroundColorTheme)
        {
            Type = SLFillType.PatternFill;
            PatternPreset = PresetPattern;
            PatternForegroundColor.SetColor(ForegroundColor, 0);
            PatternBackgroundColor.SetColor(BackgroundColorTheme, 0, 0);
        }

        /// <summary>
        ///     Set a pattern fill with a preset pattern, foreground color and background color.
        /// </summary>
        /// <param name="PresetPattern">A preset fill pattern.</param>
        /// <param name="ForegroundColor">The color to be used for the foreground.</param>
        /// <param name="BackgroundColorTheme">The theme color to be used for the background.</param>
        /// <param name="BackgroundColorTint">
        ///     The tint applied to the background theme color, ranging from -1.0 to 1.0. Negative
        ///     tints darken the theme color and positive tints lighten the theme color.
        /// </param>
        public void SetPatternFill(A.PresetPatternValues PresetPattern, Color ForegroundColor,
            SLThemeColorIndexValues BackgroundColorTheme, double BackgroundColorTint)
        {
            Type = SLFillType.PatternFill;
            PatternPreset = PresetPattern;
            PatternForegroundColor.SetColor(ForegroundColor, 0);
            PatternBackgroundColor.SetColor(BackgroundColorTheme, BackgroundColorTint, 0);
        }

        /// <summary>
        ///     Set a pattern fill with a preset pattern, foreground color and background color.
        /// </summary>
        /// <param name="PresetPattern">A preset fill pattern.</param>
        /// <param name="ForegroundColorTheme">The theme color to be used for the foreground.</param>
        /// <param name="BackgroundColor">The color to be used for the background.</param>
        public void SetPatternFill(A.PresetPatternValues PresetPattern, SLThemeColorIndexValues ForegroundColorTheme,
            Color BackgroundColor)
        {
            Type = SLFillType.PatternFill;
            PatternPreset = PresetPattern;
            PatternForegroundColor.SetColor(ForegroundColorTheme, 0, 0);
            PatternBackgroundColor.SetColor(BackgroundColor, 0);
        }

        /// <summary>
        ///     Set a pattern fill with a preset pattern, foreground color and background color.
        /// </summary>
        /// <param name="PresetPattern">A preset fill pattern.</param>
        /// <param name="ForegroundColorTheme">The theme color to be used for the foreground.</param>
        /// <param name="BackgroundColorTheme">The theme color to be used for the background.</param>
        public void SetPatternFill(A.PresetPatternValues PresetPattern, SLThemeColorIndexValues ForegroundColorTheme,
            SLThemeColorIndexValues BackgroundColorTheme)
        {
            Type = SLFillType.PatternFill;
            PatternPreset = PresetPattern;
            PatternForegroundColor.SetColor(ForegroundColorTheme, 0, 0);
            PatternBackgroundColor.SetColor(BackgroundColorTheme, 0, 0);
        }

        /// <summary>
        ///     Set a pattern fill with a preset pattern, foreground color and background color.
        /// </summary>
        /// <param name="PresetPattern">A preset fill pattern.</param>
        /// <param name="ForegroundColorTheme">The theme color to be used for the foreground.</param>
        /// <param name="BackgroundColorTheme">The theme color to be used for the background.</param>
        /// <param name="BackgroundColorTint">
        ///     The tint applied to the background theme color, ranging from -1.0 to 1.0. Negative
        ///     tints darken the theme color and positive tints lighten the theme color.
        /// </param>
        public void SetPatternFill(A.PresetPatternValues PresetPattern, SLThemeColorIndexValues ForegroundColorTheme,
            SLThemeColorIndexValues BackgroundColorTheme, double BackgroundColorTint)
        {
            Type = SLFillType.PatternFill;
            PatternPreset = PresetPattern;
            PatternForegroundColor.SetColor(ForegroundColorTheme, 0, 0);
            PatternBackgroundColor.SetColor(BackgroundColorTheme, BackgroundColorTint, 0);
        }

        /// <summary>
        ///     Set a pattern fill with a preset pattern, foreground color and background color.
        /// </summary>
        /// <param name="PresetPattern">A preset fill pattern.</param>
        /// <param name="ForegroundColorTheme">The theme color to be used for the foreground.</param>
        /// <param name="ForegroundColorTint">
        ///     The tint applied to the foreground theme color, ranging from -1.0 to 1.0. Negative
        ///     tints darken the theme color and positive tints lighten the theme color.
        /// </param>
        /// <param name="BackgroundColor">The color to be used for the background.</param>
        public void SetPatternFill(A.PresetPatternValues PresetPattern, SLThemeColorIndexValues ForegroundColorTheme,
            double ForegroundColorTint, Color BackgroundColor)
        {
            Type = SLFillType.PatternFill;
            PatternPreset = PresetPattern;
            PatternForegroundColor.SetColor(ForegroundColorTheme, ForegroundColorTint, 0);
            PatternBackgroundColor.SetColor(BackgroundColor, 0);
        }

        /// <summary>
        ///     Set a pattern fill with a preset pattern, foreground color and background color.
        /// </summary>
        /// <param name="PresetPattern">A preset fill pattern.</param>
        /// <param name="ForegroundColorTheme">The theme color to be used for the foreground.</param>
        /// <param name="ForegroundColorTint">
        ///     The tint applied to the foreground theme color, ranging from -1.0 to 1.0. Negative
        ///     tints darken the theme color and positive tints lighten the theme color.
        /// </param>
        /// <param name="BackgroundColorTheme">The theme color to be used for the background.</param>
        public void SetPatternFill(A.PresetPatternValues PresetPattern, SLThemeColorIndexValues ForegroundColorTheme,
            double ForegroundColorTint, SLThemeColorIndexValues BackgroundColorTheme)
        {
            Type = SLFillType.PatternFill;
            PatternPreset = PresetPattern;
            PatternForegroundColor.SetColor(ForegroundColorTheme, ForegroundColorTint, 0);
            PatternBackgroundColor.SetColor(BackgroundColorTheme, 0, 0);
        }

        /// <summary>
        ///     Set a pattern fill with a preset pattern, foreground color and background color.
        /// </summary>
        /// <param name="PresetPattern">A preset fill pattern.</param>
        /// <param name="ForegroundColorTheme">The theme color to be used for the foreground.</param>
        /// <param name="ForegroundColorTint">
        ///     The tint applied to the foreground theme color, ranging from -1.0 to 1.0. Negative
        ///     tints darken the theme color and positive tints lighten the theme color.
        /// </param>
        /// <param name="BackgroundColorTheme">The theme color to be used for the background.</param>
        /// <param name="BackgroundColorTint">
        ///     The tint applied to the background theme color, ranging from -1.0 to 1.0. Negative
        ///     tints darken the theme color and positive tints lighten the theme color.
        /// </param>
        public void SetPatternFill(A.PresetPatternValues PresetPattern, SLThemeColorIndexValues ForegroundColorTheme,
            double ForegroundColorTint, SLThemeColorIndexValues BackgroundColorTheme, double BackgroundColorTint)
        {
            Type = SLFillType.PatternFill;
            PatternPreset = PresetPattern;
            PatternForegroundColor.SetColor(ForegroundColorTheme, ForegroundColorTint, 0);
            PatternBackgroundColor.SetColor(BackgroundColorTheme, BackgroundColorTint, 0);
        }

        internal OpenXmlElement ToFill()
        {
            OpenXmlElement oxe = new A.NoFill();

            if (Type == SLFillType.NoFill)
                return new A.NoFill();
            if (Type == SLFillType.SolidFill)
            {
                var sf = new A.SolidFill();
                if (SolidColor.IsRgbColorModelHex)
                    sf.RgbColorModelHex = SolidColor.ToRgbColorModelHex();
                else
                    sf.SchemeColor = SolidColor.ToSchemeColor();
                return sf;
            }
            if (Type == SLFillType.GradientFill)
                return GradientColor.ToGradientFill();
            if (Type == SLFillType.BlipFill)
            {
                var bf = new A.BlipFill();
                if (BlipDpi != null) bf.Dpi = BlipDpi.Value;
                if (BlipRotateWithShape != null) bf.RotateWithShape = BlipRotateWithShape.Value;

                bf.Blip = new A.Blip();
                bf.Blip.Embed = BlipRelationshipID;
                if (BlipTransparency > 0m)
                    bf.Blip.Append(new A.AlphaModulationFixed {Amount = SLDrawingTool.CalculateAlpha(BlipTransparency)});
                bf.Append(new A.SourceRectangle());
                if (BlipTile)
                    bf.Append(new A.Tile
                    {
                        HorizontalOffset = SLDrawingTool.CalculateCoordinate(BlipOffsetX),
                        VerticalOffset = SLDrawingTool.CalculateCoordinate(BlipOffsetY),
                        HorizontalRatio = SLDrawingTool.CalculatePercentage(BlipScaleX),
                        VerticalRatio = SLDrawingTool.CalculatePercentage(BlipScaleY),
                        Flip = BlipMirrorType,
                        Alignment = BlipAlignment
                    });
                else
                    bf.Append(new A.Stretch
                    {
                        FillRectangle = new A.FillRectangle
                        {
                            Left = SLDrawingTool.CalculatePercentage(BlipLeftOffset),
                            Top = SLDrawingTool.CalculatePercentage(BlipTopOffset),
                            Right = SLDrawingTool.CalculatePercentage(BlipRightOffset),
                            Bottom = SLDrawingTool.CalculatePercentage(BlipBottomOffset)
                        }
                    });
                return bf;
            }
            if (Type == SLFillType.PatternFill)
            {
                var pf = new A.PatternFill();
                pf.Preset = A.PresetPatternValues.Trellis;

                pf.ForegroundColor = new A.ForegroundColor();
                if (PatternForegroundColor.IsRgbColorModelHex)
                    pf.ForegroundColor.RgbColorModelHex = PatternForegroundColor.ToRgbColorModelHex();
                else
                    pf.ForegroundColor.SchemeColor = PatternForegroundColor.ToSchemeColor();

                pf.BackgroundColor = new A.BackgroundColor();
                if (PatternBackgroundColor.IsRgbColorModelHex)
                    pf.BackgroundColor.RgbColorModelHex = PatternBackgroundColor.ToRgbColorModelHex();
                else
                    pf.BackgroundColor.SchemeColor = PatternBackgroundColor.ToSchemeColor();

                return pf;
            }

            return oxe;
        }

        internal SLFill Clone()
        {
            var fill = new SLFill(listThemeColors);
            fill.Type = Type;
            fill.SolidColor = SolidColor.Clone();
            fill.GradientColor = GradientColor.Clone();
            fill.BlipFileName = BlipFileName;
            fill.BlipRelationshipID = BlipRelationshipID;
            fill.BlipTile = BlipTile;
            fill.decBlipLeftOffset = decBlipLeftOffset;
            fill.decBlipRightOffset = decBlipRightOffset;
            fill.decBlipTopOffset = decBlipTopOffset;
            fill.decBlipBottomOffset = decBlipBottomOffset;
            fill.decBlipOffsetX = decBlipOffsetX;
            fill.decBlipOffsetY = decBlipOffsetY;
            fill.decBlipScaleX = decBlipScaleX;
            fill.decBlipScaleY = decBlipScaleY;
            fill.BlipAlignment = BlipAlignment;
            fill.BlipMirrorType = BlipMirrorType;
            fill.decBlipTransparency = decBlipTransparency;
            fill.BlipDpi = BlipDpi;
            fill.BlipRotateWithShape = BlipRotateWithShape;
            fill.PatternPreset = PatternPreset;
            fill.PatternForegroundColor = PatternForegroundColor.Clone();
            fill.PatternBackgroundColor = PatternBackgroundColor.Clone();

            return fill;
        }
    }
}