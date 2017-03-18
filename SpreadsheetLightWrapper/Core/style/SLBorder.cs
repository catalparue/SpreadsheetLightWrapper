using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLightWrapper.Core.misc;
using Color = System.Drawing.Color;

namespace SpreadsheetLightWrapper.Core.style
{
    /// <summary>
    ///     Encapsulates properties and methods for specifying cell borders. This simulates the
    ///     DocumentFormat.OpenXml.Spreadsheet.Border class.
    /// </summary>
    public class SLBorder
    {
        internal SLBorderProperties bpBottomBorder;
        internal SLBorderProperties bpDiagonalBorder;
        internal SLBorderProperties bpHorizontalBorder;
        internal SLBorderProperties bpLeftBorder;
        internal SLBorderProperties bpRightBorder;
        internal SLBorderProperties bpTopBorder;
        internal SLBorderProperties bpVerticalBorder;

        internal bool HasBottomBorder;

        internal bool HasDiagonalBorder;

        internal bool HasHorizontalBorder;

        internal bool HasLeftBorder;

        internal bool HasRightBorder;

        internal bool HasTopBorder;

        internal bool HasVerticalBorder;
        internal List<Color> listIndexedColors;
        internal List<Color> listThemeColors;

        /// <summary>
        ///     Initializes an instance of SLBorder. It is recommended to use CreateBorder() of the SLDocument class.
        /// </summary>
        public SLBorder()
        {
            Initialize(new List<Color>(), new List<Color>());
        }

        internal SLBorder(List<Color> ThemeColors, List<Color> IndexedColors)
        {
            Initialize(ThemeColors, IndexedColors);
        }

        /// <summary>
        ///     Encapsulates properties and methods for specifying the left border.
        /// </summary>
        public SLBorderProperties LeftBorder
        {
            get { return bpLeftBorder; }
            set
            {
                bpLeftBorder = value;
                HasLeftBorder = true;
            }
        }

        /// <summary>
        ///     Encapsulates properties and methods for specifying the right border.
        /// </summary>
        public SLBorderProperties RightBorder
        {
            get { return bpRightBorder; }
            set
            {
                bpRightBorder = value;
                HasRightBorder = true;
            }
        }

        /// <summary>
        ///     Encapsulates properties and methods for specifying the top border.
        /// </summary>
        public SLBorderProperties TopBorder
        {
            get { return bpTopBorder; }
            set
            {
                bpTopBorder = value;
                HasTopBorder = true;
            }
        }

        /// <summary>
        ///     Encapsulates properties and methods for specifying the bottom border.
        /// </summary>
        public SLBorderProperties BottomBorder
        {
            get { return bpBottomBorder; }
            set
            {
                bpBottomBorder = value;
                HasBottomBorder = true;
            }
        }

        /// <summary>
        ///     Encapsulates properties and methods for specifying the diagonal border.
        /// </summary>
        public SLBorderProperties DiagonalBorder
        {
            get { return bpDiagonalBorder; }
            set
            {
                bpDiagonalBorder = value;
                HasDiagonalBorder = true;
            }
        }

        /// <summary>
        ///     Encapsulates properties and methods for specifying the vertical border.
        /// </summary>
        public SLBorderProperties VerticalBorder
        {
            get { return bpVerticalBorder; }
            set
            {
                bpVerticalBorder = value;
                HasVerticalBorder = true;
            }
        }

        /// <summary>
        ///     Encapsulates properties and methods for specifying the horizontal border.
        /// </summary>
        public SLBorderProperties HorizontalBorder
        {
            get { return bpHorizontalBorder; }
            set
            {
                bpHorizontalBorder = value;
                HasHorizontalBorder = true;
            }
        }

        /// <summary>
        ///     Specifies if there's a diagonal line from the bottom left corner of the cell to the top right corner of the cell.
        /// </summary>
        public bool? DiagonalUp { get; set; }

        /// <summary>
        ///     Specifies if there's a diagonal line from the top left corner of the cell to the bottom right corner of the cell.
        /// </summary>
        public bool? DiagonalDown { get; set; }

        /// <summary>
        ///     Specifies if the left, right, top and bottom borders should be applied to the outside borders of a cell range.
        /// </summary>
        public bool? Outline { get; set; }

        private void Initialize(List<Color> ThemeColors, List<Color> IndexedColors)
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

        private void SetAllNull()
        {
            RemoveLeftBorder();
            RemoveRightBorder();
            RemoveTopBorder();
            RemoveBottomBorder();
            RemoveDiagonalBorder();
            RemoveVerticalBorder();
            RemoveHorizontalBorder();

            DiagonalUp = null;
            DiagonalDown = null;
            Outline = null;
        }

        /// <summary>
        ///     Remove any existing left border.
        /// </summary>
        public void RemoveLeftBorder()
        {
            bpLeftBorder = new SLBorderProperties(listThemeColors, listIndexedColors);
            HasLeftBorder = false;
        }

        /// <summary>
        ///     Remove any existing right border.
        /// </summary>
        public void RemoveRightBorder()
        {
            bpRightBorder = new SLBorderProperties(listThemeColors, listIndexedColors);
            HasRightBorder = false;
        }

        /// <summary>
        ///     Remove any existing top border.
        /// </summary>
        public void RemoveTopBorder()
        {
            bpTopBorder = new SLBorderProperties(listThemeColors, listIndexedColors);
            HasTopBorder = false;
        }

        /// <summary>
        ///     Remove any existing bottom border.
        /// </summary>
        public void RemoveBottomBorder()
        {
            bpBottomBorder = new SLBorderProperties(listThemeColors, listIndexedColors);
            HasBottomBorder = false;
        }

        /// <summary>
        ///     Remove any existing diagonal border.
        /// </summary>
        public void RemoveDiagonalBorder()
        {
            bpDiagonalBorder = new SLBorderProperties(listThemeColors, listIndexedColors);
            HasDiagonalBorder = false;
        }

        /// <summary>
        ///     Remove any existing vertical border.
        /// </summary>
        public void RemoveVerticalBorder()
        {
            bpVerticalBorder = new SLBorderProperties(listThemeColors, listIndexedColors);
            HasVerticalBorder = false;
        }

        /// <summary>
        ///     Remove any existing horizontal border.
        /// </summary>
        public void RemoveHorizontalBorder()
        {
            bpHorizontalBorder = new SLBorderProperties(listThemeColors, listIndexedColors);
            HasHorizontalBorder = false;
        }

        /// <summary>
        ///     Remove all borders.
        /// </summary>
        public void RemoveAllBorders()
        {
            RemoveLeftBorder();
            RemoveRightBorder();
            RemoveTopBorder();
            RemoveBottomBorder();
            RemoveDiagonalBorder();
            RemoveVerticalBorder();
            RemoveHorizontalBorder();
        }

        /// <summary>
        ///     Set the left border with a border style and a color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border color.</param>
        public void SetLeftBorder(BorderStyleValues BorderStyle, Color BorderColor)
        {
            LeftBorder.BorderStyle = BorderStyle;
            LeftBorder.Color = BorderColor;
        }

        /// <summary>
        ///     Set the left border with a border style and a theme color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        public void SetLeftBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor)
        {
            LeftBorder.BorderStyle = BorderStyle;
            LeftBorder.SetBorderThemeColor(BorderColor);
        }

        /// <summary>
        ///     Set the left border with a border style and a theme color, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        /// <param name="Tint">
        ///     The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color
        ///     and positive tints lighten the theme color.
        /// </param>
        public void SetLeftBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor, double Tint)
        {
            LeftBorder.BorderStyle = BorderStyle;
            LeftBorder.SetBorderThemeColor(BorderColor, Tint);
        }

        /// <summary>
        ///     Set the right border with a border style and a color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border color.</param>
        public void SetRightBorder(BorderStyleValues BorderStyle, Color BorderColor)
        {
            RightBorder.BorderStyle = BorderStyle;
            RightBorder.Color = BorderColor;
        }

        /// <summary>
        ///     Set the right border with a border style and a theme color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        public void SetRightBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor)
        {
            RightBorder.BorderStyle = BorderStyle;
            RightBorder.SetBorderThemeColor(BorderColor);
        }

        /// <summary>
        ///     Set the right border with a border style and a theme color, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        /// <param name="Tint">
        ///     The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color
        ///     and positive tints lighten the theme color.
        /// </param>
        public void SetRightBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor, double Tint)
        {
            RightBorder.BorderStyle = BorderStyle;
            RightBorder.SetBorderThemeColor(BorderColor, Tint);
        }

        /// <summary>
        ///     Set the top border with a border style and a color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border color.</param>
        public void SetTopBorder(BorderStyleValues BorderStyle, Color BorderColor)
        {
            TopBorder.BorderStyle = BorderStyle;
            TopBorder.Color = BorderColor;
        }

        /// <summary>
        ///     Set the top border with a border style and a theme color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        public void SetTopBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor)
        {
            TopBorder.BorderStyle = BorderStyle;
            TopBorder.SetBorderThemeColor(BorderColor);
        }

        /// <summary>
        ///     Set the top border with a border style and a theme color, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        /// <param name="Tint">
        ///     The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color
        ///     and positive tints lighten the theme color.
        /// </param>
        public void SetTopBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor, double Tint)
        {
            TopBorder.BorderStyle = BorderStyle;
            TopBorder.SetBorderThemeColor(BorderColor, Tint);
        }

        /// <summary>
        ///     Set the bottom border with a border style and a color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border color.</param>
        public void SetBottomBorder(BorderStyleValues BorderStyle, Color BorderColor)
        {
            BottomBorder.BorderStyle = BorderStyle;
            BottomBorder.Color = BorderColor;
        }

        /// <summary>
        ///     Set the bottom border with a border style and a theme color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        public void SetBottomBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor)
        {
            BottomBorder.BorderStyle = BorderStyle;
            BottomBorder.SetBorderThemeColor(BorderColor);
        }

        /// <summary>
        ///     Set the bottom border with a border style and a theme color, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        /// <param name="Tint">
        ///     The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color
        ///     and positive tints lighten the theme color.
        /// </param>
        public void SetBottomBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor, double Tint)
        {
            BottomBorder.BorderStyle = BorderStyle;
            BottomBorder.SetBorderThemeColor(BorderColor, Tint);
        }

        /// <summary>
        ///     Set the diagonal border with a border style and a color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border color.</param>
        public void SetDiagonalBorder(BorderStyleValues BorderStyle, Color BorderColor)
        {
            DiagonalBorder.BorderStyle = BorderStyle;
            DiagonalBorder.Color = BorderColor;
        }

        /// <summary>
        ///     Set the diagonal border with a border style and a theme color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        public void SetDiagonalBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor)
        {
            DiagonalBorder.BorderStyle = BorderStyle;
            DiagonalBorder.SetBorderThemeColor(BorderColor);
        }

        /// <summary>
        ///     Set the diagonal border with a border style and a theme color, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        /// <param name="Tint">
        ///     The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color
        ///     and positive tints lighten the theme color.
        /// </param>
        public void SetDiagonalBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor, double Tint)
        {
            DiagonalBorder.BorderStyle = BorderStyle;
            DiagonalBorder.SetBorderThemeColor(BorderColor, Tint);
        }

        /// <summary>
        ///     Set the vertical border with a border style and a color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border color.</param>
        public void SetVerticalBorder(BorderStyleValues BorderStyle, Color BorderColor)
        {
            VerticalBorder.BorderStyle = BorderStyle;
            VerticalBorder.Color = BorderColor;
        }

        /// <summary>
        ///     Set the vertical border with a border style and a theme color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        public void SetVerticalBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor)
        {
            VerticalBorder.BorderStyle = BorderStyle;
            VerticalBorder.SetBorderThemeColor(BorderColor);
        }

        /// <summary>
        ///     Set the vertical border with a border style and a theme color, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        /// <param name="Tint">
        ///     The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color
        ///     and positive tints lighten the theme color.
        /// </param>
        public void SetVerticalBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor, double Tint)
        {
            VerticalBorder.BorderStyle = BorderStyle;
            VerticalBorder.SetBorderThemeColor(BorderColor, Tint);
        }

        /// <summary>
        ///     Set the horizontal border with a border style and a color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The border color.</param>
        public void SetHorizontalBorder(BorderStyleValues BorderStyle, Color BorderColor)
        {
            HorizontalBorder.BorderStyle = BorderStyle;
            HorizontalBorder.Color = BorderColor;
        }

        /// <summary>
        ///     Set the horizontal border with a border style and a theme color.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        public void SetHorizontalBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor)
        {
            HorizontalBorder.BorderStyle = BorderStyle;
            HorizontalBorder.SetBorderThemeColor(BorderColor);
        }

        /// <summary>
        ///     Set the horizontal border with a border style and a theme color, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="BorderStyle">The border style. Default is none.</param>
        /// <param name="BorderColor">The theme color to be used.</param>
        /// <param name="Tint">
        ///     The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color
        ///     and positive tints lighten the theme color.
        /// </param>
        public void SetHorizontalBorder(BorderStyleValues BorderStyle, SLThemeColorIndexValues BorderColor, double Tint)
        {
            HorizontalBorder.BorderStyle = BorderStyle;
            HorizontalBorder.SetBorderThemeColor(BorderColor, Tint);
        }

        internal void Sync()
        {
            HasLeftBorder = LeftBorder.HasColor || LeftBorder.HasBorderStyle;
            HasRightBorder = RightBorder.HasColor || RightBorder.HasBorderStyle;
            HasTopBorder = TopBorder.HasColor || TopBorder.HasBorderStyle;
            HasBottomBorder = BottomBorder.HasColor || BottomBorder.HasBorderStyle;
            HasDiagonalBorder = DiagonalBorder.HasColor || DiagonalBorder.HasBorderStyle;
            HasVerticalBorder = VerticalBorder.HasColor || VerticalBorder.HasBorderStyle;
            HasHorizontalBorder = HorizontalBorder.HasColor || HorizontalBorder.HasBorderStyle;
        }

        /// <summary>
        ///     Form SLBorder from DocumentFormat.OpenXml.Spreadsheet.Border class.
        /// </summary>
        /// <param name="border">The source DocumentFormat.OpenXml.Spreadsheet.Border class.</param>
        public void FromBorder(Border border)
        {
            SetAllNull();

            if (border.LeftBorder != null)
            {
                HasLeftBorder = true;
                bpLeftBorder = new SLBorderProperties(listThemeColors, listIndexedColors);
                bpLeftBorder.FromBorderPropertiesType(border.LeftBorder);
            }

            if (border.RightBorder != null)
            {
                HasRightBorder = true;
                bpRightBorder = new SLBorderProperties(listThemeColors, listIndexedColors);
                bpRightBorder.FromBorderPropertiesType(border.RightBorder);
            }

            if (border.TopBorder != null)
            {
                HasTopBorder = true;
                bpTopBorder = new SLBorderProperties(listThemeColors, listIndexedColors);
                bpTopBorder.FromBorderPropertiesType(border.TopBorder);
            }

            if (border.BottomBorder != null)
            {
                HasBottomBorder = true;
                bpBottomBorder = new SLBorderProperties(listThemeColors, listIndexedColors);
                bpBottomBorder.FromBorderPropertiesType(border.BottomBorder);
            }

            if (border.DiagonalBorder != null)
            {
                HasDiagonalBorder = true;
                bpDiagonalBorder = new SLBorderProperties(listThemeColors, listIndexedColors);
                bpDiagonalBorder.FromBorderPropertiesType(border.DiagonalBorder);
            }

            if (border.VerticalBorder != null)
            {
                HasVerticalBorder = true;
                bpVerticalBorder = new SLBorderProperties(listThemeColors, listIndexedColors);
                bpVerticalBorder.FromBorderPropertiesType(border.VerticalBorder);
            }

            if (border.HorizontalBorder != null)
            {
                HasHorizontalBorder = true;
                bpHorizontalBorder = new SLBorderProperties(listThemeColors, listIndexedColors);
                bpHorizontalBorder.FromBorderPropertiesType(border.HorizontalBorder);
            }

            if (border.DiagonalUp != null) DiagonalUp = border.DiagonalUp.Value;
            else DiagonalUp = null;

            if (border.DiagonalDown != null) DiagonalDown = border.DiagonalDown.Value;
            else DiagonalDown = null;

            if (border.Outline != null) Outline = border.Outline.Value;
            else Outline = null;

            Sync();
        }

        /// <summary>
        ///     Form a DocumentFormat.OpenXml.Spreadsheet.Border class from SLBorder.
        /// </summary>
        /// <returns>A DocumentFormat.OpenXml.Spreadsheet.Border with the properties of this SLBorder class.</returns>
        public Border ToBorder()
        {
            Sync();

            var border = new Border();
            // by "default" always have left, right, top, bottom and diagonal borders, even if empty?
            border.LeftBorder = LeftBorder.ToLeftBorder();
            border.RightBorder = RightBorder.ToRightBorder();
            border.TopBorder = TopBorder.ToTopBorder();
            border.BottomBorder = BottomBorder.ToBottomBorder();
            border.DiagonalBorder = DiagonalBorder.ToDiagonalBorder();
            if (HasVerticalBorder) border.VerticalBorder = VerticalBorder.ToVerticalBorder();
            if (HasHorizontalBorder) border.HorizontalBorder = HorizontalBorder.ToHorizontalBorder();
            if (DiagonalUp != null) border.DiagonalUp = DiagonalUp.Value;
            if (DiagonalDown != null) border.DiagonalDown = DiagonalDown.Value;
            // default is true. So set property only if false
            // This reduces tag attributes
            if ((Outline != null) && !Outline.Value) border.Outline = false;

            return border;
        }

        internal void FromHash(string Hash)
        {
            var b = new Border();

            var saElementAttribute = Hash.Split(new[] {SLConstants.XmlBorderElementAttributeSeparator},
                StringSplitOptions.None);

            if (saElementAttribute.Length >= 2)
            {
                b.InnerXml = saElementAttribute[0];
                var sa = saElementAttribute[1].Split(new[] {SLConstants.XmlBorderAttributeSeparator},
                    StringSplitOptions.None);
                if (sa.Length >= 3)
                {
                    if (!sa[0].Equals("null")) b.DiagonalUp = bool.Parse(sa[0]);

                    if (!sa[1].Equals("null")) b.DiagonalDown = bool.Parse(sa[1]);

                    if (!sa[2].Equals("null")) b.Outline = bool.Parse(sa[2]);
                }
            }

            FromBorder(b);
        }

        internal string ToHash()
        {
            var b = ToBorder();
            var sXml = SLTool.RemoveNamespaceDeclaration(b.InnerXml);

            var sb = new StringBuilder();

            sb.AppendFormat("{0}{1}", sXml, SLConstants.XmlBorderElementAttributeSeparator);

            if (b.DiagonalUp != null)
                sb.AppendFormat("{0}{1}", b.DiagonalUp.Value, SLConstants.XmlBorderAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlBorderAttributeSeparator);

            if (b.DiagonalDown != null)
                sb.AppendFormat("{0}{1}", b.DiagonalDown.Value, SLConstants.XmlBorderAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlBorderAttributeSeparator);

            if (b.Outline != null) sb.AppendFormat("{0}{1}", b.Outline.Value, SLConstants.XmlBorderAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlBorderAttributeSeparator);

            return sb.ToString();
        }

        internal SLBorder Clone()
        {
            var b = new SLBorder(listThemeColors, listIndexedColors);
            b.HasLeftBorder = HasLeftBorder;
            b.bpLeftBorder = bpLeftBorder.Clone();
            b.HasRightBorder = HasRightBorder;
            b.bpRightBorder = bpRightBorder.Clone();
            b.HasTopBorder = HasTopBorder;
            b.bpTopBorder = bpTopBorder.Clone();
            b.HasBottomBorder = HasBottomBorder;
            b.bpBottomBorder = bpBottomBorder.Clone();
            b.HasDiagonalBorder = HasDiagonalBorder;
            b.bpDiagonalBorder = bpDiagonalBorder.Clone();
            b.HasVerticalBorder = HasVerticalBorder;
            b.bpVerticalBorder = bpVerticalBorder.Clone();
            b.HasHorizontalBorder = HasHorizontalBorder;
            b.bpHorizontalBorder = bpHorizontalBorder.Clone();
            b.DiagonalUp = DiagonalUp;
            b.DiagonalDown = DiagonalDown;
            b.Outline = Outline;

            return b;
        }
    }

    /// <summary>
    ///     Encapsulates properties and methods of border properties. This simulates the (abstract)
    ///     DocumentFormat.OpenXml.Spreadsheet.BorderPropertiesType class.
    /// </summary>
    public class SLBorderProperties
    {
        internal SLColor clrReal;

        internal bool HasBorderStyle;

        internal bool HasColor;
        internal List<Color> listIndexedColors;
        internal List<Color> listThemeColors;
        private BorderStyleValues vBorderStyle;

        internal SLBorderProperties(List<Color> ThemeColors, List<Color> IndexedColors)
        {
            Initialize(ThemeColors, IndexedColors);
        }

        /// <summary>
        ///     The border color.
        /// </summary>
        public Color Color
        {
            get { return clrReal.Color; }
            set
            {
                clrReal.Color = value;
                HasColor = clrReal.Color.IsEmpty ? false : true;
            }
        }

        /// <summary>
        ///     The border style. Default is none.
        /// </summary>
        public BorderStyleValues BorderStyle
        {
            get { return vBorderStyle; }
            set
            {
                vBorderStyle = value;
                HasBorderStyle = vBorderStyle != BorderStyleValues.None ? true : false;
            }
        }

        private void Initialize(List<Color> ThemeColors, List<Color> IndexedColors)
        {
            int i;
            listThemeColors = new List<Color>();
            for (i = 0; i < ThemeColors.Count; ++i)
                listThemeColors.Add(ThemeColors[i]);

            listIndexedColors = new List<Color>();
            for (i = 0; i < IndexedColors.Count; ++i)
                listIndexedColors.Add(IndexedColors[i]);

            RemoveColor();
            RemoveBorderStyle();
        }

        /// <summary>
        ///     Set the color of the border with one of the theme colors.
        /// </summary>
        /// <param name="ThemeColorIndex">The theme color to be used.</param>
        public void SetBorderThemeColor(SLThemeColorIndexValues ThemeColorIndex)
        {
            clrReal.SetThemeColor(ThemeColorIndex);
            HasColor = clrReal.Color.IsEmpty ? false : true;
        }

        /// <summary>
        ///     Set the color of the border with one of the theme colors, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="ThemeColorIndex">The theme color to be used.</param>
        /// <param name="Tint">
        ///     The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color
        ///     and positive tints lighten the theme color.
        /// </param>
        public void SetBorderThemeColor(SLThemeColorIndexValues ThemeColorIndex, double Tint)
        {
            clrReal.SetThemeColor(ThemeColorIndex, Tint);
            HasColor = clrReal.Color.IsEmpty ? false : true;
        }

        /// <summary>
        ///     Remove any existing color.
        /// </summary>
        public void RemoveColor()
        {
            clrReal = new SLColor(listThemeColors, listIndexedColors);
            HasColor = false;
        }

        /// <summary>
        ///     Remove any existing border style.
        /// </summary>
        public void RemoveBorderStyle()
        {
            vBorderStyle = BorderStyleValues.None;
            HasBorderStyle = false;
        }

        internal void FromBorderPropertiesType(LeftBorder border)
        {
            if (border.Color != null)
            {
                clrReal = new SLColor(listThemeColors, listIndexedColors);
                clrReal.FromSpreadsheetColor(border.Color);
                HasColor = !clrReal.IsEmpty();
            }
            else
            {
                RemoveColor();
            }

            if (border.Style != null) BorderStyle = border.Style.Value;
            else RemoveBorderStyle();
        }

        internal void FromBorderPropertiesType(RightBorder border)
        {
            if (border.Color != null)
            {
                clrReal = new SLColor(listThemeColors, listIndexedColors);
                clrReal.FromSpreadsheetColor(border.Color);
                HasColor = !clrReal.IsEmpty();
            }
            else
            {
                RemoveColor();
            }

            if (border.Style != null) BorderStyle = border.Style.Value;
            else RemoveBorderStyle();
        }

        internal void FromBorderPropertiesType(TopBorder border)
        {
            if (border.Color != null)
            {
                clrReal = new SLColor(listThemeColors, listIndexedColors);
                clrReal.FromSpreadsheetColor(border.Color);
                HasColor = !clrReal.IsEmpty();
            }
            else
            {
                RemoveColor();
            }

            if (border.Style != null) BorderStyle = border.Style.Value;
            else RemoveBorderStyle();
        }

        internal void FromBorderPropertiesType(BottomBorder border)
        {
            if (border.Color != null)
            {
                clrReal = new SLColor(listThemeColors, listIndexedColors);
                clrReal.FromSpreadsheetColor(border.Color);
                HasColor = !clrReal.IsEmpty();
            }
            else
            {
                RemoveColor();
            }

            if (border.Style != null) BorderStyle = border.Style.Value;
            else RemoveBorderStyle();
        }

        internal void FromBorderPropertiesType(DiagonalBorder border)
        {
            if (border.Color != null)
            {
                clrReal = new SLColor(listThemeColors, listIndexedColors);
                clrReal.FromSpreadsheetColor(border.Color);
                HasColor = !clrReal.IsEmpty();
            }
            else
            {
                RemoveColor();
            }

            if (border.Style != null) BorderStyle = border.Style.Value;
            else RemoveBorderStyle();
        }

        internal void FromBorderPropertiesType(VerticalBorder border)
        {
            if (border.Color != null)
            {
                clrReal = new SLColor(listThemeColors, listIndexedColors);
                clrReal.FromSpreadsheetColor(border.Color);
                HasColor = !clrReal.IsEmpty();
            }
            else
            {
                RemoveColor();
            }

            if (border.Style != null) BorderStyle = border.Style.Value;
            else RemoveBorderStyle();
        }

        internal void FromBorderPropertiesType(HorizontalBorder border)
        {
            if (border.Color != null)
            {
                clrReal = new SLColor(listThemeColors, listIndexedColors);
                clrReal.FromSpreadsheetColor(border.Color);
                HasColor = !clrReal.IsEmpty();
            }
            else
            {
                RemoveColor();
            }

            if (border.Style != null) BorderStyle = border.Style.Value;
            else RemoveBorderStyle();
        }

        internal LeftBorder ToLeftBorder()
        {
            var border = new LeftBorder();
            if (HasColor) border.Color = clrReal.ToSpreadsheetColor();
            if (HasBorderStyle) border.Style = BorderStyle;

            return border;
        }

        internal RightBorder ToRightBorder()
        {
            var border = new RightBorder();
            if (HasColor) border.Color = clrReal.ToSpreadsheetColor();
            if (HasBorderStyle) border.Style = BorderStyle;

            return border;
        }

        internal TopBorder ToTopBorder()
        {
            var border = new TopBorder();
            if (HasColor) border.Color = clrReal.ToSpreadsheetColor();
            if (HasBorderStyle) border.Style = BorderStyle;

            return border;
        }

        internal BottomBorder ToBottomBorder()
        {
            var border = new BottomBorder();
            if (HasColor) border.Color = clrReal.ToSpreadsheetColor();
            if (HasBorderStyle) border.Style = BorderStyle;

            return border;
        }

        internal DiagonalBorder ToDiagonalBorder()
        {
            var border = new DiagonalBorder();
            if (HasColor) border.Color = clrReal.ToSpreadsheetColor();
            if (HasBorderStyle) border.Style = BorderStyle;

            return border;
        }

        internal VerticalBorder ToVerticalBorder()
        {
            var border = new VerticalBorder();
            if (HasColor) border.Color = clrReal.ToSpreadsheetColor();
            if (HasBorderStyle) border.Style = BorderStyle;

            return border;
        }

        internal HorizontalBorder ToHorizontalBorder()
        {
            var border = new HorizontalBorder();
            if (HasColor) border.Color = clrReal.ToSpreadsheetColor();
            if (HasBorderStyle) border.Style = BorderStyle;

            return border;
        }

        internal void FromHash(string Hash)
        {
            // Just use the left border. Make sure it's consistent with the ToHash() function.
            var lb = new LeftBorder();

            var saElementAttribute = Hash.Split(new[] {SLConstants.XmlBorderPropertiesElementAttributeSeparator},
                StringSplitOptions.None);

            if (saElementAttribute.Length >= 2)
            {
                lb.InnerXml = saElementAttribute[0];
                var sa = saElementAttribute[1].Split(new[] {SLConstants.XmlBorderPropertiesAttributeSeparator},
                    StringSplitOptions.None);
                if (sa.Length >= 1)
                    if (!sa[0].Equals("null"))
                        lb.Style = (BorderStyleValues) Enum.Parse(typeof(BorderStyleValues), sa[0]);
            }

            FromBorderPropertiesType(lb);
        }

        internal string ToHash()
        {
            // Just use the left border. Make sure it's consistent with the FromHash() function.
            var lb = ToLeftBorder();
            var sXml = SLTool.RemoveNamespaceDeclaration(lb.InnerXml);

            var sb = new StringBuilder();

            sb.AppendFormat("{0}{1}", sXml, SLConstants.XmlBorderPropertiesElementAttributeSeparator);

            if (lb.Style != null)
                sb.AppendFormat("{0}{1}", lb.Style.Value, SLConstants.XmlBorderPropertiesAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlBorderPropertiesAttributeSeparator);

            return sb.ToString();
        }

        internal static string WriteToXmlTag(string BorderTag, SLBorderProperties bp)
        {
            var sb = new StringBuilder();
            sb.AppendFormat("<x:{0}", BorderTag);
            if (bp.HasBorderStyle)
                sb.AppendFormat(" style=\"{0}\"", bp.GetBorderStyleAttribute(bp.BorderStyle));

            if (bp.HasColor)
            {
                sb.Append("><x:color");
                if (bp.clrReal.Auto != null) sb.AppendFormat(" auto=\"{0}\"", bp.clrReal.Auto.Value ? "1" : "0");
                if (bp.clrReal.Indexed != null) sb.AppendFormat(" indexed=\"{0}\"", bp.clrReal.Indexed.Value);
                if (bp.clrReal.Rgb != null) sb.AppendFormat(" rgb=\"{0}\"", bp.clrReal.Rgb);
                if (bp.clrReal.Theme != null) sb.AppendFormat(" theme=\"{0}\"", bp.clrReal.Theme.Value);
                if (bp.clrReal.Tint != null) sb.AppendFormat(" tint=\"{0}\"", bp.clrReal.Tint.Value);
                sb.AppendFormat(" /></x:{0}>", BorderTag);
            }
            else
            {
                sb.Append(" />");
            }

            return sb.ToString();
        }

        internal string GetBorderStyleAttribute(BorderStyleValues bsv)
        {
            var result = "none";
            switch (bsv)
            {
                case BorderStyleValues.DashDot:
                    result = "dashDot";
                    break;
                case BorderStyleValues.DashDotDot:
                    result = "dashDotDot";
                    break;
                case BorderStyleValues.Dashed:
                    result = "dashed";
                    break;
                case BorderStyleValues.Dotted:
                    result = "dotted";
                    break;
                case BorderStyleValues.Double:
                    result = "double";
                    break;
                case BorderStyleValues.Hair:
                    result = "hair";
                    break;
                case BorderStyleValues.Medium:
                    result = "medium";
                    break;
                case BorderStyleValues.MediumDashDot:
                    result = "mediumDashDot";
                    break;
                case BorderStyleValues.MediumDashDotDot:
                    result = "mediumDashDotDot";
                    break;
                case BorderStyleValues.MediumDashed:
                    result = "mediumDashed";
                    break;
                case BorderStyleValues.None:
                    result = "none";
                    break;
                case BorderStyleValues.SlantDashDot:
                    result = "slantDashDot";
                    break;
                case BorderStyleValues.Thick:
                    result = "thick";
                    break;
                case BorderStyleValues.Thin:
                    result = "thin";
                    break;
            }

            return result;
        }

        internal SLBorderProperties Clone()
        {
            var bp = new SLBorderProperties(listThemeColors, listIndexedColors);
            bp.HasColor = HasColor;
            bp.clrReal = clrReal.Clone();
            bp.HasBorderStyle = HasBorderStyle;
            bp.vBorderStyle = vBorderStyle;

            return bp;
        }
    }
}