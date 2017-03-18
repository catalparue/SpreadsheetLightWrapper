using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLightWrapper.Core.misc;
using Color = System.Drawing.Color;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace SpreadsheetLightWrapper.Core.style
{
    /// <summary>
    ///     Encapsulates properties and methods for specifying incremental formatting. This simulates the
    ///     DocumentFormat.OpenXml.Spreadsheet.DifferentialFormat and DocumentFormat.OpenXml.Office2010.Excel.DifferentialType
    ///     classes.
    /// </summary>
    public class SLDifferentialFormat
    {
        private SLAlignment alignReal;
        private SLBorder borderReal;
        private SLFill fillReal;
        private SLFont fontReal;
        internal bool HasAlignment;

        internal bool HasBorder;

        internal bool HasFill;

        internal bool HasFont;

        internal bool HasNumberingFormat;

        internal bool HasProtection;
        internal SLNumberingFormat nfFormatCode;
        private SLProtection protectionReal;

        /// <summary>
        ///     Initializes an instance of SLDifferentialFormat.
        /// </summary>
        public SLDifferentialFormat()
        {
            SetAllNull();
        }

        /// <summary>
        ///     The alignment for incremental formatting.
        /// </summary>
        public SLAlignment Alignment
        {
            get { return alignReal; }
            set
            {
                alignReal = value;
                HasAlignment = true;
            }
        }

        /// <summary>
        ///     The protection settings for incremental formatting.
        /// </summary>
        public SLProtection Protection
        {
            get { return protectionReal; }
            set
            {
                protectionReal = value;
                HasProtection = true;
            }
        }

        /// <summary>
        ///     The numbering format for incremental formatting.
        /// </summary>
        public string FormatCode
        {
            get { return nfFormatCode.FormatCode; }
            set
            {
                nfFormatCode.FormatCode = value.Trim();
                if (nfFormatCode.FormatCode.Length > 0)
                    HasNumberingFormat = true;
                else
                    HasNumberingFormat = false;
            }
        }

        /// <summary>
        ///     The font for incremental formatting.
        /// </summary>
        public SLFont Font
        {
            get { return fontReal; }
            set
            {
                fontReal = value;
                HasFont = true;
            }
        }

        /// <summary>
        ///     The fill for incremental formatting.
        /// </summary>
        public SLFill Fill
        {
            get { return fillReal; }
            set
            {
                fillReal = value;
                HasFill = true;
            }
        }

        /// <summary>
        ///     The border for incremental formatting.
        /// </summary>
        public SLBorder Border
        {
            get { return borderReal; }
            set
            {
                borderReal = value;
                HasBorder = true;
            }
        }

        private void SetAllNull()
        {
            var listempty = new List<Color>();

            alignReal = new SLAlignment();
            HasAlignment = false;
            protectionReal = new SLProtection();
            HasProtection = false;
            nfFormatCode = new SLNumberingFormat();
            HasNumberingFormat = false;
            fontReal = new SLFont(SLConstants.OfficeThemeMajorLatinFont, SLConstants.OfficeThemeMinorLatinFont,
                listempty, listempty);
            HasFont = false;
            fillReal = new SLFill(listempty, listempty);
            HasFill = false;
            borderReal = new SLBorder(listempty, listempty);
            HasBorder = false;
        }

        internal void Sync()
        {
            HasAlignment = Alignment.HasHorizontal || Alignment.HasVertical || (Alignment.TextRotation != null) ||
                           (Alignment.WrapText != null) || (Alignment.Indent != null) ||
                           (Alignment.RelativeIndent != null) ||
                           (Alignment.JustifyLastLine != null) || (Alignment.ShrinkToFit != null) ||
                           Alignment.HasReadingOrder;
            HasProtection = (Protection.Locked != null) || (Protection.Hidden != null);
            //HasNumberingFormat
            HasFont = (Font.FontName != null) || (Font.CharacterSet != null) || (Font.FontFamily != null) ||
                      (Font.Bold != null) ||
                      (Font.Italic != null) || (Font.Strike != null) || (Font.Outline != null) || (Font.Shadow != null) ||
                      (Font.Condense != null) || (Font.Extend != null) || Font.HasFontColor || (Font.FontSize != null) ||
                      Font.HasUnderline || Font.HasVerticalAlignment || Font.HasFontScheme;
            HasFill = Fill.HasBeenAssignedValues;
            Border.Sync();
            HasBorder = Border.HasLeftBorder || Border.HasRightBorder || Border.HasTopBorder || Border.HasBottomBorder ||
                        Border.HasDiagonalBorder || Border.HasVerticalBorder || Border.HasHorizontalBorder ||
                        (Border.DiagonalUp != null) || (Border.DiagonalDown != null) || (Border.Outline != null);
        }

        internal void FromDifferentialFormat(DifferentialFormat df)
        {
            SetAllNull();

            var listempty = new List<Color>();

            if (df.Font != null)
            {
                HasFont = true;
                fontReal = new SLFont(SLConstants.OfficeThemeMajorLatinFont, SLConstants.OfficeThemeMinorLatinFont,
                    listempty, listempty);
                fontReal.FromFont(df.Font);
            }

            if (df.NumberingFormat != null)
            {
                HasNumberingFormat = true;
                nfFormatCode = new SLNumberingFormat();
                nfFormatCode.FromNumberingFormat(df.NumberingFormat);
            }

            if (df.Fill != null)
            {
                HasFill = true;
                fillReal = new SLFill(listempty, listempty);
                fillReal.FromFill(df.Fill);
            }

            if (df.Alignment != null)
            {
                HasAlignment = true;
                alignReal = new SLAlignment();
                alignReal.FromAlignment(df.Alignment);
            }

            if (df.Border != null)
            {
                HasBorder = true;
                borderReal = new SLBorder(listempty, listempty);
                borderReal.FromBorder(df.Border);
            }

            if (df.Protection != null)
            {
                HasProtection = true;
                protectionReal = new SLProtection();
                protectionReal.FromProtection(df.Protection);
            }

            Sync();
        }

        internal void FromDifferentialType(X14.DifferentialType dt)
        {
            SetAllNull();

            var listempty = new List<Color>();

            if (dt.Font != null)
            {
                HasFont = true;
                fontReal = new SLFont(SLConstants.OfficeThemeMajorLatinFont, SLConstants.OfficeThemeMinorLatinFont,
                    listempty, listempty);
                fontReal.FromFont(dt.Font);
            }

            if (dt.NumberingFormat != null)
            {
                HasNumberingFormat = true;
                nfFormatCode = new SLNumberingFormat();
                nfFormatCode.FromNumberingFormat(dt.NumberingFormat);
            }

            if (dt.Fill != null)
            {
                HasFill = true;
                fillReal = new SLFill(listempty, listempty);
                fillReal.FromFill(dt.Fill);
            }

            if (dt.Alignment != null)
            {
                HasAlignment = true;
                alignReal = new SLAlignment();
                alignReal.FromAlignment(dt.Alignment);
            }

            if (dt.Border != null)
            {
                HasBorder = true;
                borderReal = new SLBorder(listempty, listempty);
                borderReal.FromBorder(dt.Border);
            }

            if (dt.Protection != null)
            {
                HasProtection = true;
                protectionReal = new SLProtection();
                protectionReal.FromProtection(dt.Protection);
            }

            Sync();
        }

        internal DifferentialFormat ToDifferentialFormat()
        {
            Sync();

            var df = new DifferentialFormat();
            if (HasFont) df.Font = Font.ToFont();
            if (HasNumberingFormat) df.NumberingFormat = nfFormatCode.ToNumberingFormat();
            if (HasFill) df.Fill = Fill.ToFill();
            if (HasAlignment) df.Alignment = Alignment.ToAlignment();
            if (HasBorder) df.Border = Border.ToBorder();
            if (HasProtection) df.Protection = Protection.ToProtection();

            return df;
        }

        internal X14.DifferentialType ToDifferentialType()
        {
            Sync();

            var dt = new X14.DifferentialType();
            if (HasFont) dt.Font = Font.ToFont();
            if (HasNumberingFormat) dt.NumberingFormat = nfFormatCode.ToNumberingFormat();
            if (HasFill) dt.Fill = Fill.ToFill();
            if (HasAlignment) dt.Alignment = Alignment.ToAlignment();
            if (HasBorder) dt.Border = Border.ToBorder();
            if (HasProtection) dt.Protection = Protection.ToProtection();

            return dt;
        }

        internal void FromHash(string Hash)
        {
            var df = new DifferentialFormat();
            df.InnerXml = Hash;
            FromDifferentialFormat(df);
        }

        internal string ToHash()
        {
            var df = ToDifferentialFormat();
            return SLTool.RemoveNamespaceDeclaration(df.InnerXml);
        }

        internal SLDifferentialFormat Clone()
        {
            var df = new SLDifferentialFormat();
            df.HasAlignment = HasAlignment;
            df.alignReal = alignReal.Clone();
            df.HasProtection = HasProtection;
            df.protectionReal = protectionReal.Clone();
            df.HasNumberingFormat = HasNumberingFormat;
            df.nfFormatCode = nfFormatCode.Clone();
            df.HasFont = HasFont;
            df.fontReal = fontReal.Clone();
            df.HasFill = HasFill;
            df.fillReal = fillReal.Clone();
            df.HasBorder = HasBorder;
            df.borderReal = borderReal.Clone();

            return df;
        }
    }
}