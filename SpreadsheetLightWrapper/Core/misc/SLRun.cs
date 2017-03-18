using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLightWrapper.Core.style;
using Color = System.Drawing.Color;

namespace SpreadsheetLightWrapper.Core.misc
{
    /// <summary>
    ///     Encapsulates properties and methods for rich text runs. This simulates the DocumentFormat.OpenXml.Spreadsheet.Run
    ///     class.
    /// </summary>
    public class SLRun
    {
        /// <summary>
        ///     Initializes an instance of SLRun.
        /// </summary>
        public SLRun()
        {
            SetAllNull();
        }

        /// <summary>
        ///     The font styles.
        /// </summary>
        public SLFont Font { get; set; }

        /// <summary>
        ///     The text.
        /// </summary>
        public string Text { get; set; }

        private void SetAllNull()
        {
            Font = new SLFont(SLConstants.OfficeThemeMajorLatinFont, SLConstants.OfficeThemeMinorLatinFont,
                new List<Color>(), new List<Color>());
            Text = string.Empty;
        }

        internal void FromRun(Run r)
        {
            SetAllNull();

            using (var oxr = OpenXmlReader.Create(r))
            {
                while (oxr.Read())
                    if (oxr.ElementType == typeof(Text))
                    {
                        Text = ((Text) oxr.LoadCurrentElement()).Text;
                    }
                    else if (oxr.ElementType == typeof(RunFont))
                    {
                        var rft = (RunFont) oxr.LoadCurrentElement();
                        if (rft.Val != null) Font.FontName = rft.Val.Value;
                    }
                    else if (oxr.ElementType == typeof(RunPropertyCharSet))
                    {
                        var rpcs = (RunPropertyCharSet) oxr.LoadCurrentElement();
                        if (rpcs.Val != null) Font.CharacterSet = rpcs.Val.Value;
                    }
                    else if (oxr.ElementType == typeof(FontFamily))
                    {
                        var ff = (FontFamily) oxr.LoadCurrentElement();
                        if (ff.Val != null) Font.FontFamily = ff.Val.Value;
                    }
                    else if (oxr.ElementType == typeof(Bold))
                    {
                        var b = (Bold) oxr.LoadCurrentElement();
                        if (b.Val != null) Font.Bold = b.Val.Value;
                        else Font.Bold = true;
                    }
                    else if (oxr.ElementType == typeof(Italic))
                    {
                        var itlc = (Italic) oxr.LoadCurrentElement();
                        if (itlc.Val != null) Font.Italic = itlc.Val.Value;
                        else Font.Italic = true;
                    }
                    else if (oxr.ElementType == typeof(Strike))
                    {
                        var strk = (Strike) oxr.LoadCurrentElement();
                        if (strk.Val != null) Font.Strike = strk.Val.Value;
                        else Font.Strike = true;
                    }
                    else if (oxr.ElementType == typeof(Outline))
                    {
                        var outln = (Outline) oxr.LoadCurrentElement();
                        if (outln.Val != null) Font.Outline = outln.Val.Value;
                        else Font.Outline = true;
                    }
                    else if (oxr.ElementType == typeof(Shadow))
                    {
                        var shdw = (Shadow) oxr.LoadCurrentElement();
                        if (shdw.Val != null) Font.Shadow = shdw.Val.Value;
                        else Font.Shadow = true;
                    }
                    else if (oxr.ElementType == typeof(Condense))
                    {
                        var cdns = (Condense) oxr.LoadCurrentElement();
                        if (cdns.Val != null) Font.Condense = cdns.Val.Value;
                        else Font.Condense = true;
                    }
                    else if (oxr.ElementType == typeof(Extend))
                    {
                        var ext = (Extend) oxr.LoadCurrentElement();
                        if (ext.Val != null) Font.Extend = ext.Val.Value;
                        else Font.Extend = true;
                    }
                    else if (oxr.ElementType == typeof(DocumentFormat.OpenXml.Spreadsheet.Color))
                    {
                        Font.clrFontColor.FromSpreadsheetColor(
                            (DocumentFormat.OpenXml.Spreadsheet.Color) oxr.LoadCurrentElement());
                        Font.HasFontColor = !Font.clrFontColor.IsEmpty();
                    }
                    else if (oxr.ElementType == typeof(FontSize))
                    {
                        var ftsz = (FontSize) oxr.LoadCurrentElement();
                        if (ftsz.Val != null) Font.FontSize = ftsz.Val.Value;
                    }
                    else if (oxr.ElementType == typeof(Underline))
                    {
                        var und = (Underline) oxr.LoadCurrentElement();
                        if (und.Val != null) Font.Underline = und.Val.Value;
                        else Font.Underline = UnderlineValues.Single;
                    }
                    else if (oxr.ElementType == typeof(VerticalTextAlignment))
                    {
                        var vta = (VerticalTextAlignment) oxr.LoadCurrentElement();
                        if (vta.Val != null) Font.VerticalAlignment = vta.Val.Value;
                    }
                    else if (oxr.ElementType == typeof(FontScheme))
                    {
                        var ftsch = (FontScheme) oxr.LoadCurrentElement();
                        if (ftsch.Val != null) Font.FontScheme = ftsch.Val.Value;
                    }
            }
        }

        internal Run ToRun()
        {
            var r = new Run();
            r.RunProperties = new RunProperties();

            if (Font.FontName != null)
                r.RunProperties.Append(new RunFont {Val = Font.FontName});

            if (Font.CharacterSet != null)
                r.RunProperties.Append(new RunPropertyCharSet {Val = Font.CharacterSet.Value});

            if (Font.FontFamily != null)
                r.RunProperties.Append(new FontFamily {Val = Font.FontFamily.Value});

            if ((Font.Bold != null) && Font.Bold.Value)
                r.RunProperties.Append(new Bold());

            if ((Font.Italic != null) && Font.Italic.Value)
                r.RunProperties.Append(new Italic());

            if ((Font.Strike != null) && Font.Strike.Value)
                r.RunProperties.Append(new Strike());

            if ((Font.Outline != null) && Font.Outline.Value)
                r.RunProperties.Append(new Outline());

            if ((Font.Shadow != null) && Font.Shadow.Value)
                r.RunProperties.Append(new Shadow());

            if ((Font.Condense != null) && Font.Condense.Value)
                r.RunProperties.Append(new Condense());

            if ((Font.Extend != null) && Font.Extend.Value)
                r.RunProperties.Append(new Extend());

            if (Font.HasFontColor)
                r.RunProperties.Append(Font.clrFontColor.ToSpreadsheetColor());

            if (Font.FontSize != null)
                r.RunProperties.Append(new FontSize {Val = Font.FontSize.Value});

            if (Font.HasUnderline)
                r.RunProperties.Append(new Underline {Val = Font.Underline});

            if (Font.HasVerticalAlignment)
                r.RunProperties.Append(new VerticalTextAlignment {Val = Font.VerticalAlignment});

            if (Font.HasFontScheme)
                r.RunProperties.Append(new FontScheme {Val = Font.FontScheme});

            r.Text = new Text(Text);
            if (SLTool.ToPreserveSpace(Text)) r.Text.Space = SpaceProcessingModeValues.Preserve;

            return r;
        }

        /// <summary>
        ///     Clone a new instance of SLRun.
        /// </summary>
        /// <returns>An SLRun object.</returns>
        public SLRun Clone()
        {
            var r = new SLRun();
            r.Font = Font.Clone();
            r.Text = Text;

            return r;
        }
    }
}