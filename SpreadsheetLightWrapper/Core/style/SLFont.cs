using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLightWrapper.Core.Drawing;
using SpreadsheetLightWrapper.Core.misc;
using A = DocumentFormat.OpenXml.Drawing;
using Color = System.Drawing.Color;
using FontScheme = DocumentFormat.OpenXml.Spreadsheet.FontScheme;
using Outline = DocumentFormat.OpenXml.Spreadsheet.Outline;
using Underline = DocumentFormat.OpenXml.Spreadsheet.Underline;

namespace SpreadsheetLightWrapper.Core.style
{
    /// <summary>
    ///     Encapsulates properties and methods for fonts. This simulates the DocumentFormat.OpenXml.Spreadsheet.Font class.
    /// </summary>
    public class SLFont
    {
        internal SLColor clrFontColor;

        internal bool HasFontColor;

        internal bool HasFontScheme;

        internal bool HasUnderline;

        internal bool HasVerticalAlignment;
        internal List<Color> listIndexedColors;
        internal List<Color> listThemeColors;
        private FontSchemeValues vFontScheme;
        private UnderlineValues vUnderline;
        private VerticalAlignmentRunValues vVerticalAlignment;

        /// <summary>
        ///     Initializes an instance of SLFont. It is recommended to use CreateFont() of the SLDocument class.
        /// </summary>
        public SLFont()
        {
            Initialize(SLConstants.OfficeThemeMajorLatinFont, SLConstants.OfficeThemeMinorLatinFont, new List<Color>(),
                new List<Color>());
        }

        internal SLFont(string MajorFont, string MinorFont, List<Color> ThemeColors, List<Color> IndexedColors)
        {
            Initialize(MajorFont, MinorFont, ThemeColors, IndexedColors);
        }

        internal string MajorFont { get; set; }
        internal string MinorFont { get; set; }

        /// <summary>
        ///     SheetName of the font.
        /// </summary>
        public string FontName { get; set; }

        /// <summary>
        ///     The font character set of the font text. It is recommended not to explicitly set this property. This is used when
        ///     the given font name is not available on the computer, and a suitable alternative font is used. The character set
        ///     value is operating system dependent. Possible value (not exhaustive): 0 - ANSI_CHARSET, 1 - DEFAULT_CHARSET, 2 -
        ///     SYMBOL_CHARSET.
        /// </summary>
        public int? CharacterSet { get; set; }

        /// <summary>
        ///     The font family of the font text. It is recommended not to explicitly set this property. Values as follows (might
        ///     not be exhaustive): 0 - Not applicable, 1 - Roman, 2 - Swiss, 3 - Modern, 4 - Script, 5 - Decorative.
        /// </summary>
        public int? FontFamily { get; set; }

        /// <summary>
        ///     Specifies if the font text should be in bold.
        /// </summary>
        public bool? Bold { get; set; }

        /// <summary>
        ///     Specifies if the font text should be in italic.
        /// </summary>
        public bool? Italic { get; set; }

        /// <summary>
        ///     Specifies if the font text should have a strikethrough.
        /// </summary>
        public bool? Strike { get; set; }

        /// <summary>
        ///     Specifies if the inner and outer borders of each character of the font text should be displayed. This makes the
        ///     font text appear as if in bold.
        /// </summary>
        public bool? Outline { get; set; }

        /// <summary>
        ///     Specifies if there's a shadow behind and at the bottom-right of the font text. It is a Macintosh compatibility
        ///     setting.
        ///     It is recommended not to use this property because SpreadsheetML applications are not required to use this
        ///     property.
        /// </summary>
        public bool? Shadow { get; set; }

        /// <summary>
        ///     Specifies if the font text should be squeezed together. It is a Macintosh compatibility setting.
        ///     It is recommended not to use this property because SpreadsheetML applications are not required to use this
        ///     property.
        /// </summary>
        public bool? Condense { get; set; }

        /// <summary>
        ///     Specifies if the font text should be stretched out. It is a legacy spreadsheet compatibility setting.
        ///     It is recommended not to use this property because SpreadsheetML applications are not required to use this
        ///     property.
        /// </summary>
        public bool? Extend { get; set; }

        /// <summary>
        ///     The color of the font text.
        /// </summary>
        public Color FontColor
        {
            get { return clrFontColor.Color; }
            set
            {
                clrFontColor.Color = value;
                HasFontColor = clrFontColor.Color.IsEmpty ? false : true;
            }
        }

        /// <summary>
        ///     The size of the font text in points (1 point is 1/72 of an inch).
        /// </summary>
        public double? FontSize { get; set; }

        // default is single, but for hashing we use none as default
        /// <summary>
        ///     Specifies the underline formatting style of the font text.
        /// </summary>
        public UnderlineValues Underline
        {
            get { return vUnderline; }
            set
            {
                vUnderline = value;
                HasUnderline = vUnderline != UnderlineValues.None ? true : false;
            }
        }

        /// <summary>
        ///     Specifies the vertical position of the font text.
        /// </summary>
        public VerticalAlignmentRunValues VerticalAlignment
        {
            get { return vVerticalAlignment; }
            set
            {
                vVerticalAlignment = value;
                HasVerticalAlignment = true;
            }
        }

        /// <summary>
        ///     Specifies the font scheme. Used particularly as part of a theme definition. A major font scheme is usually used for
        ///     heading text. A minor font scheme is used for body text.
        /// </summary>
        public FontSchemeValues FontScheme
        {
            get { return vFontScheme; }
            set
            {
                vFontScheme = value;
                HasFontScheme = true;
            }
        }

        private void Initialize(string MajorFont, string MinorFont, List<Color> ThemeColors, List<Color> IndexedColors)
        {
            this.MajorFont = MajorFont;
            this.MinorFont = MinorFont;

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
            FontName = null;
            CharacterSet = null;
            FontFamily = null;
            Bold = null;
            Italic = null;
            Strike = null;
            Outline = null;
            Shadow = null;
            Condense = null;
            Extend = null;
            clrFontColor = new SLColor(listThemeColors, listIndexedColors);
            HasFontColor = false;
            FontSize = null;
            vUnderline = UnderlineValues.None;
            HasUnderline = false;
            vVerticalAlignment = VerticalAlignmentRunValues.Baseline;
            HasVerticalAlignment = false;
            vFontScheme = FontSchemeValues.None;
            HasFontScheme = false;
        }

        /// <summary>
        ///     Set the font, given a font name and font size.
        /// </summary>
        /// <param name="FontName">The name of the font to be used.</param>
        /// <param name="FontSize">The size of the font in points.</param>
        public void SetFont(string FontName, double FontSize)
        {
            this.FontName = FontName;
            this.FontSize = FontSize;
            CharacterSet = null;
            FontFamily = null;
            vFontScheme = FontSchemeValues.None;
            HasFontScheme = false;
        }

        /// <summary>
        ///     Set the font, given a font scheme and font size.
        /// </summary>
        /// <param name="FontScheme">
        ///     The font scheme. If None is given, the current theme's minor font will be used (but if the
        ///     theme is changed, the text remains as of the old theme's minor font instead of the new theme's minor font).
        /// </param>
        /// <param name="FontSize">The size of the font in points.</param>
        public void SetFont(FontSchemeValues FontScheme, double FontSize)
        {
            switch (FontScheme)
            {
                case FontSchemeValues.Major:
                    FontName = MajorFont;
                    this.FontScheme = FontSchemeValues.Major;
                    break;
                case FontSchemeValues.Minor:
                    FontName = MinorFont;
                    this.FontScheme = FontSchemeValues.Minor;
                    break;
                case FontSchemeValues.None:
                    FontName = MinorFont;
                    vFontScheme = FontSchemeValues.None;
                    HasFontScheme = false;
                    break;
            }
            this.FontSize = FontSize;
            CharacterSet = null;
            FontFamily = null;
        }

        /// <summary>
        ///     Set the font color with one of the theme colors.
        /// </summary>
        /// <param name="ThemeColorIndex">The theme color to be used.</param>
        public void SetFontThemeColor(SLThemeColorIndexValues ThemeColorIndex)
        {
            clrFontColor.SetThemeColor(ThemeColorIndex);
            HasFontColor = clrFontColor.Color.IsEmpty ? false : true;
        }

        /// <summary>
        ///     Set the font color with one of the theme colors, modifying the theme color with a tint value.
        /// </summary>
        /// <param name="ThemeColorIndex">The theme color to be used.</param>
        /// <param name="Tint">
        ///     The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color
        ///     and positive tints lighten the theme color.
        /// </param>
        public void SetFontThemeColor(SLThemeColorIndexValues ThemeColorIndex, double Tint)
        {
            clrFontColor.SetThemeColor(ThemeColorIndex, Tint);
            HasFontColor = clrFontColor.Color.IsEmpty ? false : true;
        }

        internal void FromFont(Font f)
        {
            SetAllNull();

            if ((f.FontName != null) && (f.FontName.Val != null))
                FontName = f.FontName.Val.Value;

            if ((f.FontCharSet != null) && (f.FontCharSet.Val != null))
                CharacterSet = f.FontCharSet.Val.Value;

            if ((f.FontFamilyNumbering != null) && (f.FontFamilyNumbering.Val != null))
                FontFamily = f.FontFamilyNumbering.Val.Value;

            if (f.Bold != null)
                if (f.Bold.Val == null) Bold = true;
                else if (f.Bold.Val.Value) Bold = true;

            if (f.Italic != null)
                if (f.Italic.Val == null) Italic = true;
                else if (f.Italic.Val.Value) Italic = true;

            if (f.Strike != null)
                if (f.Strike.Val == null) Strike = true;
                else if (f.Strike.Val.Value) Strike = true;

            if (f.Outline != null)
                if (f.Outline.Val == null) Outline = true;
                else if (f.Outline.Val.Value) Outline = true;

            if (f.Shadow != null)
                if (f.Shadow.Val == null) Shadow = true;
                else if (f.Shadow.Val.Value) Shadow = true;

            if (f.Condense != null)
                if (f.Condense.Val == null) Condense = true;
                else if (f.Condense.Val.Value) Condense = true;

            if (f.Extend != null)
                if (f.Extend.Val == null) Extend = true;
                else if (f.Extend.Val.Value) Extend = true;

            if (f.Color != null)
            {
                clrFontColor = new SLColor(listThemeColors, listIndexedColors);
                clrFontColor.FromSpreadsheetColor(f.Color);
                HasFontColor = !clrFontColor.IsEmpty();
            }

            if ((f.FontSize != null) && (f.FontSize.Val != null))
                FontSize = f.FontSize.Val.Value;

            if (f.Underline != null)
                if (f.Underline.Val != null)
                    Underline = f.Underline.Val.Value;
                else
                    Underline = UnderlineValues.Single;

            if ((f.VerticalTextAlignment != null) && (f.VerticalTextAlignment.Val != null))
                VerticalAlignment = f.VerticalTextAlignment.Val.Value;

            if ((f.FontScheme != null) && (f.FontScheme.Val != null))
                FontScheme = f.FontScheme.Val.Value;
        }

        internal Font ToFont()
        {
            var f = new Font();
            if (FontName != null) f.FontName = new FontName {Val = FontName};
            if (CharacterSet != null) f.FontCharSet = new FontCharSet {Val = CharacterSet.Value};
            if (FontFamily != null) f.FontFamilyNumbering = new FontFamilyNumbering {Val = FontFamily.Value};
            if ((Bold != null) && Bold.Value) f.Bold = new Bold();
            if ((Italic != null) && Italic.Value) f.Italic = new Italic();
            if ((Strike != null) && Strike.Value) f.Strike = new Strike();
            if ((Outline != null) && Outline.Value) f.Outline = new Outline();
            if ((Shadow != null) && Shadow.Value) f.Shadow = new Shadow();
            if ((Condense != null) && Condense.Value) f.Condense = new Condense();
            if ((Extend != null) && Extend.Value) f.Extend = new Extend();
            if (HasFontColor) f.Color = clrFontColor.ToSpreadsheetColor();
            if (FontSize != null) f.FontSize = new FontSize {Val = FontSize.Value};
            if (HasUnderline)
                if (Underline == UnderlineValues.Single)
                    f.Underline = new Underline();
                else
                    f.Underline = new Underline {Val = Underline};
            if (HasVerticalAlignment) f.VerticalTextAlignment = new VerticalTextAlignment {Val = VerticalAlignment};
            if (HasFontScheme) f.FontScheme = new FontScheme {Val = FontScheme};

            return f;
        }

        internal void FromHash(string Hash)
        {
            var font = new Font();
            font.InnerXml = Hash;
            FromFont(font);
        }

        internal string ToHash()
        {
            var font = ToFont();
            return SLTool.RemoveNamespaceDeclaration(font.InnerXml);
        }

        // SLFont takes on extra duties so you don't have to learn more classes. Just like SLRstType.
        internal A.Paragraph ToParagraph()
        {
            var para = new A.Paragraph();
            para.ParagraphProperties = new A.ParagraphProperties();

            var defrunprops = new A.DefaultRunProperties();

            var sFont = string.Empty;
            if ((FontName != null) && (FontName.Length > 0)) sFont = FontName;

            if (HasFontScheme)
                if (FontScheme == FontSchemeValues.Major) sFont = "+mj-lt";
                else if (FontScheme == FontSchemeValues.Minor) sFont = "+mn-lt";

            if (HasFontColor)
            {
                var clr = new SLColorTransform(new List<Color>());
                if ((clrFontColor.Rgb != null) && (clrFontColor.Rgb.Length > 0))
                {
                    clr.SetColor(SLTool.ToColor(clrFontColor.Rgb), 0);

                    defrunprops.Append(new A.SolidFill
                    {
                        RgbColorModelHex = clr.ToRgbColorModelHex()
                    });
                }
                else if (clrFontColor.Theme != null)
                {
                    // potential casting error? If the SLFont class was set properly, there shouldn't be errors...
                    var themeindex = (SLThemeColorIndexValues) clrFontColor.Theme.Value;
                    if (clrFontColor.Tint != null)
                        clr.SetColor(themeindex, clrFontColor.Tint.Value, 0);
                    else
                        clr.SetColor(themeindex, 0, 0);

                    defrunprops.Append(new A.SolidFill
                    {
                        SchemeColor = clr.ToSchemeColor()
                    });
                }
            }

            if (sFont.Length > 0) defrunprops.Append(new A.LatinFont {Typeface = sFont});

            if (FontSize != null) defrunprops.FontSize = (int) (FontSize.Value*100);

            if (Bold != null) defrunprops.Bold = Bold.Value;

            if (Italic != null) defrunprops.Italic = Italic.Value;

            if (HasUnderline)
                if ((Underline == UnderlineValues.Single) || (Underline == UnderlineValues.SingleAccounting))
                    defrunprops.Underline = A.TextUnderlineValues.Single;
                else if ((Underline == UnderlineValues.Double) || (Underline == UnderlineValues.DoubleAccounting))
                    defrunprops.Underline = A.TextUnderlineValues.Double;

            if (Strike != null)
                defrunprops.Strike = Strike.Value ? A.TextStrikeValues.SingleStrike : A.TextStrikeValues.NoStrike;

            if (HasVerticalAlignment)
                if (VerticalAlignment == VerticalAlignmentRunValues.Superscript)
                    defrunprops.Baseline = 30000;
                else if (VerticalAlignment == VerticalAlignmentRunValues.Subscript)
                    defrunprops.Baseline = -25000;
                else
                    defrunprops.Baseline = 0;

            para.ParagraphProperties.Append(defrunprops);

            return para;
        }

        /// <summary>
        ///     Clone a new instance of SLFont with identical font settings.
        /// </summary>
        /// <returns>An SLFont object with identical font settings.</returns>
        public SLFont Clone()
        {
            var font = new SLFont(MajorFont, MinorFont, listThemeColors, listIndexedColors);

            font.FontName = FontName;
            font.CharacterSet = CharacterSet;
            font.FontFamily = FontFamily;
            font.Bold = Bold;
            font.Italic = Italic;
            font.Strike = Strike;
            font.Outline = Outline;
            font.Shadow = Shadow;
            font.Condense = Condense;
            font.Extend = Extend;
            font.clrFontColor = clrFontColor.Clone();
            font.HasFontColor = HasFontColor;
            font.FontSize = FontSize;
            font.vUnderline = vUnderline;
            font.HasUnderline = HasUnderline;
            font.vVerticalAlignment = vVerticalAlignment;
            font.HasVerticalAlignment = HasVerticalAlignment;
            font.vFontScheme = vFontScheme;
            font.HasFontScheme = HasFontScheme;

            return font;
        }
    }
}