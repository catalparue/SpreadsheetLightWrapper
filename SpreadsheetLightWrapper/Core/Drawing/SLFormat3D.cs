using System.Collections.Generic;
using System.Drawing;
using SpreadsheetLightWrapper.Core.misc;
using A = DocumentFormat.OpenXml.Drawing;

namespace SpreadsheetLightWrapper.Core.Drawing
{
    /// <summary>
    ///     Encapsulates 3D shape properties. Works together with SLRotation3D class.
    ///     This simulates some properties of DocumentFormat.OpenXml.Drawing.Scene3DType
    ///     and DocumentFormat.OpenXml.Drawing.Shape3DType classes. The reason for this mixing
    ///     is because Excel separates different properties from both classes into 2 separate sections
    ///     on the user interface (3-D Format and 3-D Rotation). Hence SLRotation3D and SLFormat3D
    ///     classes instead of straightforward mapping of the SDK Scene3DType and Shape3DType classes.
    /// </summary>
    public class SLFormat3D
    {
        internal bool bHasLighting;
        internal SLColorTransform clrContourColor;
        internal SLColorTransform clrExtrusionColor;

        internal decimal decAngle;

        internal decimal decBevelBottomHeight;

        internal decimal decBevelBottomWidth;

        internal decimal decBevelTopHeight;

        internal decimal decBevelTopWidth;

        internal decimal decContourWidth;

        internal decimal decExtrusionHeight;

        internal bool HasContourColor;

        internal bool HasExtrusionColor;
        internal List<Color> listThemeColors;

        internal A.BevelPresetValues vBevelBottomPreset;

        internal A.BevelPresetValues vBevelTopPreset;

        internal A.LightRigValues vLighting;

        internal SLFormat3D(List<Color> ThemeColors)
        {
            int i;
            listThemeColors = new List<Color>();
            for (i = 0; i < ThemeColors.Count; ++i)
                listThemeColors.Add(ThemeColors[i]);

            SetAllNull();
        }

        /// <summary>
        ///     Specifies if there's a top bevel. This is read-only.
        /// </summary>
        public bool HasBevelTop { get; private set; }

        /// <summary>
        ///     The bevel type of the top bevel. Default is circle.
        /// </summary>
        public A.BevelPresetValues BevelTopPreset
        {
            get { return vBevelTopPreset; }
            set
            {
                vBevelTopPreset = value;
                HasBevelTop = true;
            }
        }

        /// <summary>
        ///     Width of the top bevel, ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt. Accurate to
        ///     1/12700 of a point.
        /// </summary>
        public decimal BevelTopWidth
        {
            get { return decBevelTopWidth; }
            set
            {
                decBevelTopWidth = value;
                if (decBevelTopWidth < 0m) decBevelTopWidth = 0m;
                if (decBevelTopWidth > 2147483647m) decBevelTopWidth = 2147483647m;
                HasBevelTop = true;
            }
        }

        /// <summary>
        ///     Height of the top bevel, ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt. Accurate to
        ///     1/12700 of a point.
        /// </summary>
        public decimal BevelTopHeight
        {
            get { return decBevelTopHeight; }
            set
            {
                decBevelTopHeight = value;
                if (decBevelTopHeight < 0m) decBevelTopHeight = 0m;
                if (decBevelTopHeight > 2147483647m) decBevelTopHeight = 2147483647m;
                HasBevelTop = true;
            }
        }

        /// <summary>
        ///     Specifies if there's a bottom bevel. This is read-only.
        /// </summary>
        public bool HasBevelBottom { get; private set; }

        /// <summary>
        ///     The bevel type of the bottom bevel. Default is circle.
        /// </summary>
        public A.BevelPresetValues BevelBottomPreset
        {
            get { return vBevelBottomPreset; }
            set
            {
                vBevelBottomPreset = value;
                HasBevelBottom = true;
            }
        }

        /// <summary>
        ///     Width of the bottom bevel, ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt. Accurate to
        ///     1/12700 of a point.
        /// </summary>
        public decimal BevelBottomWidth
        {
            get { return decBevelBottomWidth; }
            set
            {
                decBevelBottomWidth = value;
                if (decBevelBottomWidth < 0m) decBevelBottomWidth = 0m;
                if (decBevelBottomWidth > 2147483647m) decBevelBottomWidth = 2147483647m;
                HasBevelBottom = true;
            }
        }

        /// <summary>
        ///     Height of the bottom bevel, ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt. Accurate
        ///     to 1/12700 of a point.
        /// </summary>
        public decimal BevelBottomHeight
        {
            get { return decBevelBottomHeight; }
            set
            {
                decBevelBottomHeight = value;
                if (decBevelBottomHeight < 0m) decBevelBottomHeight = 0m;
                if (decBevelBottomHeight > 2147483647m) decBevelBottomHeight = 2147483647m;
                HasBevelBottom = true;
            }
        }

        /// <summary>
        ///     The extrusion color, also known as the depth color. This is read-only.
        /// </summary>
        public Color ExtrusionColor
        {
            get { return clrExtrusionColor.DisplayColor; }
        }

        /// <summary>
        ///     Extrusion height, ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt. Accurate to 1/12700
        ///     of a point.
        ///     The Microsoft Excel user interface uses the term "Depth".
        /// </summary>
        public decimal ExtrusionHeight
        {
            get { return decExtrusionHeight; }
            set
            {
                decExtrusionHeight = value;
                if (decExtrusionHeight < 0m) decExtrusionHeight = 0m;
                if (decExtrusionHeight > 2147483647m) decExtrusionHeight = 2147483647m;
            }
        }

        /// <summary>
        ///     The contour color. This is read-only.
        /// </summary>
        public Color ContourColor
        {
            get { return clrContourColor.DisplayColor; }
        }

        /// <summary>
        ///     Contour width, ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt. Accurate to 1/12700 of
        ///     a point.
        ///     The Microsoft Excel user interface uses the term "Size".
        /// </summary>
        public decimal ContourWidth
        {
            get { return decContourWidth; }
            set
            {
                decContourWidth = value;
                if (decContourWidth < 0m) decContourWidth = 0m;
                if (decContourWidth > 2147483647m) decContourWidth = 2147483647m;
            }
        }

        /// <summary>
        ///     The preset material used. Default is WarmMatte.
        /// </summary>
        public A.PresetMaterialTypeValues Material { get; set; }

        /// <summary>
        ///     Specifies if there's lighting.
        /// </summary>
        public bool HasLighting
        {
            get { return bHasLighting; }
        }

        /// <summary>
        ///     The type of lighting used.
        /// </summary>
        public A.LightRigValues Lighting
        {
            get { return vLighting; }
            set
            {
                vLighting = value;
                bHasLighting = true;
            }
        }

        /// <summary>
        ///     Angle of the lighting, ranging from 0 degrees to 359.9 degrees. This is set only when <see cref="Lighting" /> is
        ///     also set.
        /// </summary>
        public decimal Angle
        {
            get { return decAngle; }
            set
            {
                decAngle = value;
                if (decAngle < 0m) decAngle = 0m;
                if (decAngle >= 360m) decAngle = 359.9m;
            }
        }

        private void SetAllNull()
        {
            SetNoBevelTop();
            SetNoBevelBottom();
            SetNoExtrusion();
            SetNoContour();

            Material = A.PresetMaterialTypeValues.WarmMatte;

            SetNoLighting();
        }

        /// <summary>
        ///     Set the top bevel.
        /// </summary>
        /// <param name="BevelPreset">The bevel type.</param>
        /// <param name="Width">
        ///     Bevel width ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt. Accurate
        ///     to 1/12700 of a point.
        /// </param>
        /// <param name="Height">
        ///     Bevel height ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt. Accurate
        ///     to 1/12700 of a point.
        /// </param>
        public void SetBevelTop(A.BevelPresetValues BevelPreset, decimal Width, decimal Height)
        {
            vBevelTopPreset = BevelPreset;
            BevelTopWidth = Width;
            BevelTopHeight = Height;
            HasBevelTop = true;
        }

        /// <summary>
        ///     Remove the top bevel.
        /// </summary>
        public void SetNoBevelTop()
        {
            vBevelTopPreset = A.BevelPresetValues.Circle;
            decBevelTopWidth = 6;
            decBevelTopHeight = 6;
            HasBevelTop = false;
        }

        /// <summary>
        ///     Set the bottom bevel.
        /// </summary>
        /// <param name="BevelPreset">The bevel type.</param>
        /// <param name="Width">
        ///     Bevel width ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt. Accurate
        ///     to 1/12700 of a point.
        /// </param>
        /// <param name="Height">
        ///     Bevel height ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt. Accurate
        ///     to 1/12700 of a point.
        /// </param>
        public void SetBevelBottom(A.BevelPresetValues BevelPreset, decimal Width, decimal Height)
        {
            vBevelBottomPreset = BevelPreset;
            BevelBottomWidth = Width;
            BevelBottomHeight = Height;
            HasBevelBottom = true;
        }

        /// <summary>
        ///     Remove the bottom bevel.
        /// </summary>
        public void SetNoBevelBottom()
        {
            vBevelBottomPreset = A.BevelPresetValues.Circle;
            decBevelBottomWidth = 6;
            decBevelBottomHeight = 6;
            HasBevelBottom = false;
        }

        /// <summary>
        ///     Remove any extrusion (or depth) settings.
        /// </summary>
        public void SetNoExtrusion()
        {
            clrExtrusionColor = new SLColorTransform(listThemeColors);
            HasExtrusionColor = false;
            decExtrusionHeight = 0;
        }

        /// <summary>
        ///     Set the extrusion (or depth) color.
        /// </summary>
        /// <param name="Color">The color to be used.</param>
        public void SetExtrusionColor(Color Color)
        {
            if (!Color.IsEmpty)
            {
                clrExtrusionColor.SetColor(Color, 0);
                HasExtrusionColor = true;
            }
        }

        /// <summary>
        ///     Set the extrusion (or depth) color.
        /// </summary>
        /// <param name="Color">The theme color to be used.</param>
        /// <param name="Tint">
        ///     The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color
        ///     and positive tints lighten the theme color.
        /// </param>
        public void SetExtrusionColor(SLThemeColorIndexValues Color, double Tint)
        {
            clrExtrusionColor.SetColor(Color, Tint, 0);
            HasExtrusionColor = true;
        }

        /// <summary>
        ///     Set the extrusion.
        /// </summary>
        /// <param name="Color">The color to be used.</param>
        /// <param name="Height">
        ///     Extrusion height, ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt.
        ///     Accurate to 1/12700 of a point.
        /// </param>
        public void SetExtrusion(Color Color, decimal Height)
        {
            if (!Color.IsEmpty)
            {
                clrExtrusionColor.SetColor(Color, 0);
                HasExtrusionColor = true;
            }
            ExtrusionHeight = Height;
        }

        /// <summary>
        ///     Set the extrusion.
        /// </summary>
        /// <param name="Color">The theme color to be used.</param>
        /// <param name="Tint">
        ///     The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color
        ///     and positive tints lighten the theme color.
        /// </param>
        /// <param name="Height">
        ///     Extrusion height, ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt.
        ///     Accurate to 1/12700 of a point.
        /// </param>
        public void SetExtrusion(SLThemeColorIndexValues Color, double Tint, decimal Height)
        {
            clrExtrusionColor.SetColor(Color, Tint, 0);
            HasExtrusionColor = true;
            ExtrusionHeight = Height;
        }

        /// <summary>
        ///     Remove any contour settings.
        /// </summary>
        public void SetNoContour()
        {
            clrContourColor = new SLColorTransform(listThemeColors);
            HasContourColor = false;
            decContourWidth = 0;
        }

        /// <summary>
        ///     Set the contour color.
        /// </summary>
        /// <param name="Color">The color to be used.</param>
        public void SetContourColor(Color Color)
        {
            if (!Color.IsEmpty)
            {
                clrContourColor.SetColor(Color, 0);
                HasContourColor = true;
            }
        }

        /// <summary>
        ///     Set the contour color.
        /// </summary>
        /// <param name="Color">The theme color to be used.</param>
        /// <param name="Tint">
        ///     The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color
        ///     and positive tints lighten the theme color.
        /// </param>
        public void SetContourColor(SLThemeColorIndexValues Color, double Tint)
        {
            clrContourColor.SetColor(Color, Tint, 0);
            HasContourColor = true;
        }

        /// <summary>
        ///     Set the contour.
        /// </summary>
        /// <param name="Color">The color to be used.</param>
        /// <param name="Width">
        ///     Contour width, ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt.
        ///     Accurate to 1/12700 of a point.
        /// </param>
        public void SetContour(Color Color, decimal Width)
        {
            if (!Color.IsEmpty)
            {
                clrContourColor.SetColor(Color, 0);
                HasContourColor = true;
            }
            ContourWidth = Width;
        }

        /// <summary>
        ///     Set the contour.
        /// </summary>
        /// <param name="Color">The theme color to be used.</param>
        /// <param name="Tint">
        ///     The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color
        ///     and positive tints lighten the theme color.
        /// </param>
        /// <param name="Width">
        ///     Contour width, ranging from 0 pt to 2147483647 pt. However, a suggested maximum is 1584 pt.
        ///     Accurate to 1/12700 of a point.
        /// </param>
        public void SetContour(SLThemeColorIndexValues Color, double Tint, decimal Width)
        {
            clrContourColor.SetColor(Color, Tint, 0);
            HasContourColor = true;
            ContourWidth = Width;
        }

        /// <summary>
        ///     Remove any lighting settings.
        /// </summary>
        public void SetNoLighting()
        {
            vLighting = A.LightRigValues.ThreePoints;
            bHasLighting = false;
            decAngle = 0;
        }

        internal SLFormat3D Clone()
        {
            var format = new SLFormat3D(listThemeColors);
            format.HasBevelTop = HasBevelTop;
            format.vBevelTopPreset = vBevelTopPreset;
            format.decBevelTopWidth = decBevelTopWidth;
            format.decBevelTopHeight = decBevelTopHeight;
            format.HasBevelBottom = HasBevelBottom;
            format.vBevelBottomPreset = vBevelBottomPreset;
            format.decBevelBottomWidth = decBevelBottomWidth;
            format.decBevelBottomHeight = decBevelBottomHeight;
            format.HasExtrusionColor = HasExtrusionColor;
            format.clrExtrusionColor = clrExtrusionColor.Clone();
            format.decExtrusionHeight = decExtrusionHeight;
            format.HasContourColor = HasContourColor;
            format.clrContourColor = clrContourColor.Clone();
            format.decContourWidth = decContourWidth;
            format.Material = Material;
            format.bHasLighting = bHasLighting;
            format.vLighting = vLighting;
            format.decAngle = decAngle;

            return format;
        }
    }
}