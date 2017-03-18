using System.Collections.Generic;
using System.Drawing;
using SpreadsheetLightWrapper.Core.misc;
using A = DocumentFormat.OpenXml.Drawing;

namespace SpreadsheetLightWrapper.Core.Drawing
{
    /// <summary>
    ///     Encapsulates properties and methods for specifying shadow effects.
    ///     This simulates the DocumentFormat.OpenXml.Drawing.InnerShadow and DocumentFormat.OpenXml.Drawing.OuterShadow
    ///     classes.
    /// </summary>
    public class SLShadowEffect
    {
        internal List<Color> listThemeColors;

        internal SLShadowEffect(List<Color> ThemeColors)
        {
            int i;
            listThemeColors = new List<Color>();
            for (i = 0; i < ThemeColors.Count; ++i)
                listThemeColors.Add(ThemeColors[i]);

            SetAllNull();
        }

        /// <summary>
        ///     doubles as HasShadow variable
        /// </summary>
        internal bool? IsInnerShadow { get; set; }

        internal SLColorTransform InnerShadowColor { get; set; }
        internal decimal InnerShadowBlurRadius { get; set; }
        internal decimal InnerShadowDistance { get; set; }
        internal decimal InnerShadowDirection { get; set; }

        internal SLColorTransform OuterShadowColor { get; set; }
        internal decimal OuterShadowBlurRadius { get; set; }
        internal decimal OuterShadowDistance { get; set; }
        internal decimal OuterShadowDirection { get; set; }
        internal decimal OuterShadowHorizontalRatio { get; set; }
        internal decimal OuterShadowVerticalRatio { get; set; }
        internal decimal OuterShadowHorizontalSkew { get; set; }
        internal decimal OuterShadowVerticalSkew { get; set; }
        internal A.RectangleAlignmentValues OuterShadowAlignment { get; set; }
        internal bool OuterShadowRotateWithShape { get; set; }

        /// <summary>
        ///     The shadow color.
        /// </summary>
        public Color Color
        {
            get
            {
                if (IsInnerShadow != null)
                {
                    if (IsInnerShadow.Value)
                        return InnerShadowColor.DisplayColor;
                    return OuterShadowColor.DisplayColor;
                }
                return new Color();
            }
        }

        /// <summary>
        ///     Transparency of the shadow color ranging from 0% to 100%. Accurate to 1/1000 of a percent.
        /// </summary>
        public decimal Transparency
        {
            get
            {
                if (IsInnerShadow != null)
                {
                    if (IsInnerShadow.Value)
                        return InnerShadowColor.Transparency;
                    return OuterShadowColor.Transparency;
                }
                return 0;
            }
            set
            {
                if (IsInnerShadow != null)
                    if (IsInnerShadow.Value)
                        InnerShadowColor.Transparency = value;
                    else
                        OuterShadowColor.Transparency = value;
            }
        }

        /// <summary>
        ///     Specifies the size of the shadow in percentage. While there's no restriction in range, consider a range of 1% to
        ///     200%. Accurate to 1/1000th of a percent.
        /// </summary>
        public decimal Size
        {
            get { return OuterShadowHorizontalRatio; }
            set
            {
                var dec = value;
                OuterShadowHorizontalRatio = dec;
                OuterShadowVerticalRatio = dec;
            }
        }

        /// <summary>
        ///     Shadow blur, ranging from 0 pt to 2147483647 pt (but consider a maximum of 100 pt). Accurate to 1/12700 of a point.
        /// </summary>
        public decimal Blur
        {
            get
            {
                if (IsInnerShadow != null)
                {
                    if (IsInnerShadow.Value)
                        return InnerShadowBlurRadius;
                    return OuterShadowBlurRadius;
                }
                return 0;
            }
            set
            {
                if (IsInnerShadow != null)
                {
                    var dec = value;
                    if (dec < 0m) dec = 0m;
                    if (dec > 100m) dec = 100m;

                    if (IsInnerShadow.Value)
                        InnerShadowBlurRadius = dec;
                    else
                        OuterShadowBlurRadius = dec;
                }
            }
        }

        /// <summary>
        ///     Angle of shadow projection, ranging from 0 degrees to 359.9 degrees. 0 degrees means the shadow is to the right of
        ///     the picture, 90 degrees means it's below, 180 degrees means it's to the left and 270 degrees means it's above.
        ///     Accurate to 1/60000 of a degree. Default value is 0 degrees.
        /// </summary>
        public decimal Angle
        {
            get
            {
                if (IsInnerShadow != null)
                {
                    if (IsInnerShadow.Value)
                        return InnerShadowDirection;
                    return OuterShadowDirection;
                }
                return 0;
            }
            set
            {
                if (IsInnerShadow != null)
                {
                    var dec = value;
                    if (dec < 0m) dec = 0m;
                    if (dec >= 360m) dec = 359.9m;

                    if (IsInnerShadow.Value)
                        InnerShadowDirection = dec;
                    else
                        InnerShadowDirection = dec;
                }
            }
        }

        /// <summary>
        ///     Distance of shadow away from source object, ranging from 0 pt to 2147483647 pt (but consider a maximum of 200 pt).
        ///     Accurate to 1/12700 of a point. Default value is 0 pt.
        /// </summary>
        public decimal Distance
        {
            get
            {
                if (IsInnerShadow != null)
                {
                    if (IsInnerShadow.Value)
                        return InnerShadowDistance;
                    return OuterShadowDistance;
                }
                return 0;
            }
            set
            {
                if (IsInnerShadow != null)
                {
                    var dec = value;
                    if (dec < 0m) dec = 0m;
                    if (dec > 200m) dec = 200m;

                    if (IsInnerShadow.Value)
                        InnerShadowDistance = dec;
                    else
                        OuterShadowDistance = dec;
                }
            }
        }

        private void SetAllNull()
        {
            IsInnerShadow = null;

            InnerShadowColor = new SLColorTransform(listThemeColors);
            InnerShadowBlurRadius = 0;
            InnerShadowDistance = 0;
            InnerShadowDirection = 0;

            OuterShadowColor = new SLColorTransform(listThemeColors);
            OuterShadowBlurRadius = 0;
            OuterShadowDistance = 0;
            OuterShadowDirection = 0;
            OuterShadowHorizontalRatio = 100;
            OuterShadowVerticalRatio = 100;
            OuterShadowHorizontalSkew = 0;
            OuterShadowVerticalSkew = 0;
            OuterShadowAlignment = A.RectangleAlignmentValues.Bottom;
            OuterShadowRotateWithShape = true;
        }

        /// <summary>
        ///     Set the shadow color.
        /// </summary>
        /// <param name="Color">The color to be used.</param>
        /// <param name="Transparency">Transparency of the color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        public void SetShadowColor(Color Color, decimal Transparency)
        {
            if (IsInnerShadow != null)
                if (IsInnerShadow.Value)
                    InnerShadowColor.SetColor(Color, Transparency);
                else
                    OuterShadowColor.SetColor(Color, Transparency);
        }

        /// <summary>
        ///     Set the shadow color.
        /// </summary>
        /// <param name="Color">The theme color used.</param>
        /// <param name="Tint">
        ///     The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color
        ///     and positive tints lighten the theme color.
        /// </param>
        /// <param name="Transparency">Transparency of the color ranging from 0% to 100%. Accurate to 1/1000 of a percent.</param>
        public void SetShadowColor(SLThemeColorIndexValues Color, double Tint, decimal Transparency)
        {
            if (IsInnerShadow != null)
                if (IsInnerShadow.Value)
                    InnerShadowColor.SetColor(Color, Tint, Transparency);
                else
                    OuterShadowColor.SetColor(Color, Tint, Transparency);
        }

        /// <summary>
        ///     Set a shadow using a preset.
        /// </summary>
        /// <param name="Preset">The preset to be used.</param>
        public void SetPreset(SLShadowPresetValues Preset)
        {
            var clr = Color.FromArgb(0, 0, 0);

            switch (Preset)
            {
                case SLShadowPresetValues.None:
                    SetAllNull();
                    break;
                case SLShadowPresetValues.OuterDiagonalBottomRight:
                    IsInnerShadow = false;
                    OuterShadowColor.SetColor(clr, 60);
                    OuterShadowHorizontalRatio = 100;
                    OuterShadowVerticalRatio = 100;
                    OuterShadowHorizontalSkew = 0;
                    OuterShadowVerticalSkew = 0;
                    OuterShadowBlurRadius = 4;
                    OuterShadowDirection = 45;
                    OuterShadowDistance = 3;
                    OuterShadowAlignment = A.RectangleAlignmentValues.TopLeft;
                    OuterShadowRotateWithShape = false;
                    break;
                case SLShadowPresetValues.OuterBottom:
                    IsInnerShadow = false;
                    OuterShadowColor.SetColor(clr, 60);
                    OuterShadowHorizontalRatio = 100;
                    OuterShadowVerticalRatio = 100;
                    OuterShadowHorizontalSkew = 0;
                    OuterShadowVerticalSkew = 0;
                    OuterShadowBlurRadius = 4;
                    OuterShadowDirection = 90;
                    OuterShadowDistance = 3;
                    OuterShadowAlignment = A.RectangleAlignmentValues.Top;
                    OuterShadowRotateWithShape = false;
                    break;
                case SLShadowPresetValues.OuterDiagonalBottomLeft:
                    IsInnerShadow = false;
                    OuterShadowColor.SetColor(clr, 60);
                    OuterShadowHorizontalRatio = 100;
                    OuterShadowVerticalRatio = 100;
                    OuterShadowHorizontalSkew = 0;
                    OuterShadowVerticalSkew = 0;
                    OuterShadowBlurRadius = 4;
                    OuterShadowDirection = 135;
                    OuterShadowDistance = 3;
                    OuterShadowAlignment = A.RectangleAlignmentValues.TopRight;
                    OuterShadowRotateWithShape = false;
                    break;
                case SLShadowPresetValues.OuterRight:
                    IsInnerShadow = false;
                    OuterShadowColor.SetColor(clr, 60);
                    OuterShadowHorizontalRatio = 100;
                    OuterShadowVerticalRatio = 100;
                    OuterShadowHorizontalSkew = 0;
                    OuterShadowVerticalSkew = 0;
                    OuterShadowBlurRadius = 4;
                    OuterShadowDirection = 0;
                    OuterShadowDistance = 3;
                    OuterShadowAlignment = A.RectangleAlignmentValues.Left;
                    OuterShadowRotateWithShape = false;
                    break;
                case SLShadowPresetValues.OuterCenter:
                    IsInnerShadow = false;
                    OuterShadowColor.SetColor(clr, 60);
                    OuterShadowHorizontalRatio = 102;
                    OuterShadowVerticalRatio = 102;
                    OuterShadowHorizontalSkew = 0;
                    OuterShadowVerticalSkew = 0;
                    OuterShadowBlurRadius = 5;
                    OuterShadowDirection = 0;
                    OuterShadowDistance = 0;
                    OuterShadowAlignment = A.RectangleAlignmentValues.Center;
                    OuterShadowRotateWithShape = false;
                    break;
                case SLShadowPresetValues.OuterLeft:
                    IsInnerShadow = false;
                    OuterShadowColor.SetColor(clr, 60);
                    OuterShadowHorizontalRatio = 100;
                    OuterShadowVerticalRatio = 100;
                    OuterShadowHorizontalSkew = 0;
                    OuterShadowVerticalSkew = 0;
                    OuterShadowBlurRadius = 4;
                    OuterShadowDirection = 180;
                    OuterShadowDistance = 3;
                    OuterShadowAlignment = A.RectangleAlignmentValues.Right;
                    OuterShadowRotateWithShape = false;
                    break;
                case SLShadowPresetValues.OuterDiagonalTopRight:
                    IsInnerShadow = false;
                    OuterShadowColor.SetColor(clr, 60);
                    OuterShadowHorizontalRatio = 100;
                    OuterShadowVerticalRatio = 100;
                    OuterShadowHorizontalSkew = 0;
                    OuterShadowVerticalSkew = 0;
                    OuterShadowBlurRadius = 4;
                    OuterShadowDirection = 315;
                    OuterShadowDistance = 3;
                    OuterShadowAlignment = A.RectangleAlignmentValues.BottomLeft;
                    OuterShadowRotateWithShape = false;
                    break;
                case SLShadowPresetValues.OuterTop:
                    IsInnerShadow = false;
                    OuterShadowColor.SetColor(clr, 60);
                    OuterShadowHorizontalRatio = 100;
                    OuterShadowVerticalRatio = 100;
                    OuterShadowHorizontalSkew = 0;
                    OuterShadowVerticalSkew = 0;
                    OuterShadowBlurRadius = 4;
                    OuterShadowDirection = 270;
                    OuterShadowDistance = 3;
                    OuterShadowAlignment = A.RectangleAlignmentValues.Bottom;
                    OuterShadowRotateWithShape = false;
                    break;
                case SLShadowPresetValues.OuterDiagonalTopLeft:
                    IsInnerShadow = false;
                    OuterShadowColor.SetColor(clr, 60);
                    OuterShadowHorizontalRatio = 100;
                    OuterShadowVerticalRatio = 100;
                    OuterShadowHorizontalSkew = 0;
                    OuterShadowVerticalSkew = 0;
                    OuterShadowBlurRadius = 4;
                    OuterShadowDirection = 225;
                    OuterShadowDistance = 3;
                    OuterShadowAlignment = A.RectangleAlignmentValues.BottomRight;
                    OuterShadowRotateWithShape = false;
                    break;
                case SLShadowPresetValues.InnerDiagonalTopLeft:
                    IsInnerShadow = true;
                    InnerShadowColor.SetColor(clr, 50);
                    InnerShadowBlurRadius = 5;
                    InnerShadowDirection = 225;
                    InnerShadowDistance = 4;
                    break;
                case SLShadowPresetValues.InnerTop:
                    IsInnerShadow = true;
                    InnerShadowColor.SetColor(clr, 50);
                    InnerShadowBlurRadius = 5;
                    InnerShadowDirection = 270;
                    InnerShadowDistance = 4;
                    break;
                case SLShadowPresetValues.InnerDiagonalTopRight:
                    IsInnerShadow = true;
                    InnerShadowColor.SetColor(clr, 50);
                    InnerShadowBlurRadius = 5;
                    InnerShadowDirection = 315;
                    InnerShadowDistance = 4;
                    break;
                case SLShadowPresetValues.InnerLeft:
                    IsInnerShadow = true;
                    InnerShadowColor.SetColor(clr, 50);
                    InnerShadowBlurRadius = 5;
                    InnerShadowDirection = 180;
                    InnerShadowDistance = 4;
                    break;
                case SLShadowPresetValues.InnerCenter:
                    IsInnerShadow = true;
                    InnerShadowColor.SetColor(clr, 0);
                    InnerShadowBlurRadius = 9;
                    InnerShadowDirection = 0;
                    InnerShadowDistance = 0;
                    break;
                case SLShadowPresetValues.InnerRight:
                    IsInnerShadow = true;
                    InnerShadowColor.SetColor(clr, 50);
                    InnerShadowBlurRadius = 5;
                    InnerShadowDirection = 0;
                    InnerShadowDistance = 4;
                    break;
                case SLShadowPresetValues.InnerDiagonalBottomLeft:
                    IsInnerShadow = true;
                    InnerShadowColor.SetColor(clr, 50);
                    InnerShadowBlurRadius = 5;
                    InnerShadowDirection = 135;
                    InnerShadowDistance = 4;
                    break;
                case SLShadowPresetValues.InnerBottom:
                    IsInnerShadow = true;
                    InnerShadowColor.SetColor(clr, 50);
                    InnerShadowBlurRadius = 5;
                    InnerShadowDirection = 90;
                    InnerShadowDistance = 4;
                    break;
                case SLShadowPresetValues.InnerDiagonalBottomRight:
                    IsInnerShadow = true;
                    InnerShadowColor.SetColor(clr, 50);
                    InnerShadowBlurRadius = 5;
                    InnerShadowDirection = 45;
                    InnerShadowDistance = 4;
                    break;
                case SLShadowPresetValues.PerspectiveDiagonalUpperLeft:
                    IsInnerShadow = false;
                    OuterShadowColor.SetColor(clr, 80);
                    OuterShadowHorizontalRatio = 100;
                    OuterShadowVerticalRatio = 23;
                    OuterShadowHorizontalSkew = 20;
                    OuterShadowVerticalSkew = 0;
                    OuterShadowBlurRadius = 6;
                    OuterShadowDirection = 225;
                    OuterShadowDistance = 0;
                    OuterShadowAlignment = A.RectangleAlignmentValues.BottomRight;
                    OuterShadowRotateWithShape = false;
                    break;
                case SLShadowPresetValues.PerspectiveDiagonalUpperRight:
                    IsInnerShadow = false;
                    OuterShadowColor.SetColor(clr, 80);
                    OuterShadowHorizontalRatio = 100;
                    OuterShadowVerticalRatio = 23;
                    OuterShadowHorizontalSkew = -20;
                    OuterShadowVerticalSkew = 0;
                    OuterShadowBlurRadius = 6;
                    OuterShadowDirection = 315;
                    OuterShadowDistance = 0;
                    OuterShadowAlignment = A.RectangleAlignmentValues.BottomLeft;
                    OuterShadowRotateWithShape = false;
                    break;
                case SLShadowPresetValues.PerspectiveBelow:
                    IsInnerShadow = false;
                    OuterShadowColor.SetColor(clr, 85);
                    OuterShadowHorizontalRatio = 90;
                    OuterShadowVerticalRatio = 100;
                    OuterShadowHorizontalSkew = 0;
                    OuterShadowVerticalSkew = -0.3166667m;
                    OuterShadowBlurRadius = 12;
                    OuterShadowDirection = 90;
                    OuterShadowDistance = 25;
                    OuterShadowAlignment = A.RectangleAlignmentValues.Bottom;
                    OuterShadowRotateWithShape = false;
                    break;
                case SLShadowPresetValues.PerspectiveDiagonalLowerLeft:
                    IsInnerShadow = false;
                    OuterShadowColor.SetColor(clr, 80);
                    OuterShadowHorizontalRatio = 100;
                    OuterShadowVerticalRatio = -23;
                    OuterShadowHorizontalSkew = 13.34m;
                    OuterShadowVerticalSkew = 0;
                    OuterShadowBlurRadius = 6;
                    OuterShadowDirection = 135;
                    OuterShadowDistance = 1;
                    OuterShadowAlignment = A.RectangleAlignmentValues.BottomRight;
                    OuterShadowRotateWithShape = false;
                    break;
                case SLShadowPresetValues.PerspectiveDiagonalLowerRight:
                    IsInnerShadow = false;
                    OuterShadowColor.SetColor(clr, 80);
                    OuterShadowHorizontalRatio = 100;
                    OuterShadowVerticalRatio = -23;
                    OuterShadowHorizontalSkew = -13.34m;
                    OuterShadowVerticalSkew = 0;
                    OuterShadowBlurRadius = 6;
                    OuterShadowDirection = 45;
                    OuterShadowDistance = 1;
                    OuterShadowAlignment = A.RectangleAlignmentValues.BottomLeft;
                    OuterShadowRotateWithShape = false;
                    break;
            }
        }

        // TODO overload setting of inner and outer shadow functions here

        internal A.InnerShadow ToInnerShadow()
        {
            var ishad = new A.InnerShadow();
            if (InnerShadowColor.IsRgbColorModelHex)
                ishad.RgbColorModelHex = InnerShadowColor.ToRgbColorModelHex();
            else
                ishad.SchemeColor = InnerShadowColor.ToSchemeColor();

            if (InnerShadowBlurRadius != 0)
                ishad.BlurRadius = SLDrawingTool.CalculatePositiveCoordinate(InnerShadowBlurRadius);

            if (InnerShadowDistance != 0)
                ishad.Distance = SLDrawingTool.CalculatePositiveCoordinate(InnerShadowDistance);

            if (InnerShadowDirection != 0)
                ishad.Direction = SLDrawingTool.CalculatePositiveFixedAngle(InnerShadowDirection);

            return ishad;
        }

        internal A.OuterShadow ToOuterShadow()
        {
            var os = new A.OuterShadow();

            if (OuterShadowColor.IsRgbColorModelHex)
                os.RgbColorModelHex = OuterShadowColor.ToRgbColorModelHex();
            else
                os.SchemeColor = OuterShadowColor.ToSchemeColor();

            if (OuterShadowBlurRadius != 0)
                os.BlurRadius = SLDrawingTool.CalculatePositiveCoordinate(OuterShadowBlurRadius);

            if (OuterShadowDistance != 0)
                os.Distance = SLDrawingTool.CalculatePositiveCoordinate(OuterShadowDistance);

            if (OuterShadowDirection != 0)
                os.Direction = SLDrawingTool.CalculatePositiveFixedAngle(OuterShadowDirection);

            if (OuterShadowHorizontalRatio != 100m)
                os.HorizontalRatio = SLDrawingTool.CalculatePercentage(OuterShadowHorizontalRatio);

            if (OuterShadowVerticalRatio != 100m)
                os.VerticalRatio = SLDrawingTool.CalculatePercentage(OuterShadowVerticalRatio);

            if (OuterShadowHorizontalSkew != 0m)
                os.HorizontalSkew = SLDrawingTool.CalculateFixedAngle(OuterShadowHorizontalSkew);

            if (OuterShadowVerticalSkew != 0m)
                os.VerticalSkew = SLDrawingTool.CalculateFixedAngle(OuterShadowVerticalSkew);

            if (OuterShadowAlignment != A.RectangleAlignmentValues.Bottom) os.Alignment = OuterShadowAlignment;

            if (!OuterShadowRotateWithShape) os.RotateWithShape = OuterShadowRotateWithShape;

            return os;
        }

        internal SLShadowEffect Clone()
        {
            var se = new SLShadowEffect(listThemeColors);
            se.IsInnerShadow = IsInnerShadow;
            se.InnerShadowColor = InnerShadowColor.Clone();
            se.InnerShadowBlurRadius = InnerShadowBlurRadius;
            se.InnerShadowDistance = InnerShadowDistance;
            se.InnerShadowDirection = InnerShadowDirection;
            se.OuterShadowColor = OuterShadowColor.Clone();
            se.OuterShadowBlurRadius = OuterShadowBlurRadius;
            se.OuterShadowDistance = OuterShadowDistance;
            se.OuterShadowDirection = OuterShadowDirection;
            se.OuterShadowHorizontalRatio = OuterShadowHorizontalRatio;
            se.OuterShadowVerticalRatio = OuterShadowVerticalRatio;
            se.OuterShadowHorizontalSkew = OuterShadowHorizontalSkew;
            se.OuterShadowVerticalSkew = OuterShadowVerticalSkew;
            se.OuterShadowAlignment = OuterShadowAlignment;
            se.OuterShadowRotateWithShape = OuterShadowRotateWithShape;

            return se;
        }
    }
}