using System;
using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;
using Ups.Toolkit.SpreadsheetLight.Core.misc;

namespace Ups.Toolkit.SpreadsheetLight.Core.style
{
    /// <summary>
    ///     Specifies reading order.
    /// </summary>
    public enum SLAlignmentReadingOrderValues
    {
        /// <summary>
        ///     Reading order is context dependent.
        /// </summary>
        ContextDependent = 0,

        /// <summary>
        ///     Reading order is from left to right.
        /// </summary>
        LeftToRight = 1,

        /// <summary>
        ///     Reading order is from right to left.
        /// </summary>
        RightToLeft = 2
    }

    /// <summary>
    ///     Encapsulates properties and methods for text alignment in cells. This simulates the
    ///     DocumentFormat.OpenXml.Spreadsheet.Alignment class.
    /// </summary>
    public class SLAlignment
    {
        internal bool HasHorizontal;

        internal bool HasReadingOrder;

        internal bool HasVertical;

        private int? iTextRotation;
        private HorizontalAlignmentValues vHorizontal;
        private SLAlignmentReadingOrderValues vReadingOrder;
        private VerticalAlignmentValues vVertical;

        /// <summary>
        ///     Initializes an instance of SLAlignment.
        /// </summary>
        public SLAlignment()
        {
            SetAllNull();
        }

        /// <summary>
        ///     Specifies the horizontal alignment. Default value is General.
        /// </summary>
        public HorizontalAlignmentValues Horizontal
        {
            get { return vHorizontal; }
            set
            {
                vHorizontal = value;
                HasHorizontal = true;
            }
        }

        /// <summary>
        ///     Specifies the vertical alignment. Default value is Bottom.
        /// </summary>
        public VerticalAlignmentValues Vertical
        {
            get { return vVertical; }
            set
            {
                vVertical = value;
                HasVertical = true;
            }
        }

        /// <summary>
        ///     Specifies the rotation angle of the text, ranging from -90 degrees to 90 degrees. Default value is 0 degrees.
        /// </summary>
        public int? TextRotation
        {
            get { return iTextRotation; }
            set
            {
                if ((value >= -90) && (value <= 90))
                    iTextRotation = value;
                else
                    iTextRotation = null;
            }
        }

        /// <summary>
        ///     Specifies if the text in the cell should be wrapped.
        /// </summary>
        public bool? WrapText { get; set; }

        /// <summary>
        ///     Specifies the indent. Each unit value equals 3 spaces.
        /// </summary>
        public uint? Indent { get; set; }

        /// <summary>
        ///     This property is used when the class is part of a SLDifferentialFormat class. It specifies the indent value in
        ///     addition to the given Indent property.
        /// </summary>
        public int? RelativeIndent { get; set; }

        /// <summary>
        ///     Specifies if the last line should be justified (usually for East Asian fonts).
        /// </summary>
        public bool? JustifyLastLine { get; set; }

        /// <summary>
        ///     Specifies if the text in the cell should be shrunk to fit the cell.
        /// </summary>
        public bool? ShrinkToFit { get; set; }

        /// <summary>
        ///     Specifies the reading order of the text in the cell.
        /// </summary>
        public SLAlignmentReadingOrderValues ReadingOrder
        {
            get { return vReadingOrder; }
            set
            {
                vReadingOrder = value;
                HasReadingOrder = true;
            }
        }

        private void SetAllNull()
        {
            vHorizontal = HorizontalAlignmentValues.General;
            HasHorizontal = false;
            vVertical = VerticalAlignmentValues.Bottom;
            HasVertical = false;
            TextRotation = null;
            WrapText = null;
            Indent = null;
            RelativeIndent = null;
            JustifyLastLine = null;
            ShrinkToFit = null;
            vReadingOrder = SLAlignmentReadingOrderValues.LeftToRight;
            HasReadingOrder = false;
        }

        internal void FromAlignment(Alignment align)
        {
            SetAllNull();

            if (align.Horizontal != null) Horizontal = align.Horizontal.Value;
            if (align.Vertical != null) Vertical = align.Vertical.Value;

            if ((align.TextRotation != null) && (align.TextRotation.Value <= 180))
                TextRotation = TextRotationToIntuitiveValue(align.TextRotation.Value);

            if (align.WrapText != null) WrapText = align.WrapText.Value;
            if (align.Indent != null) Indent = align.Indent.Value;
            if (align.RelativeIndent != null) RelativeIndent = align.RelativeIndent.Value;
            if (align.JustifyLastLine != null) JustifyLastLine = align.JustifyLastLine.Value;
            if (align.ShrinkToFit != null) ShrinkToFit = align.ShrinkToFit.Value;

            if (align.ReadingOrder != null)
                switch (align.ReadingOrder.Value)
                {
                    case (uint) SLAlignmentReadingOrderValues.ContextDependent:
                        ReadingOrder = SLAlignmentReadingOrderValues.ContextDependent;
                        break;
                    case (uint) SLAlignmentReadingOrderValues.LeftToRight:
                        ReadingOrder = SLAlignmentReadingOrderValues.LeftToRight;
                        break;
                    case (uint) SLAlignmentReadingOrderValues.RightToLeft:
                        ReadingOrder = SLAlignmentReadingOrderValues.RightToLeft;
                        break;
                }
        }

        internal Alignment ToAlignment()
        {
            var align = new Alignment();
            if (HasHorizontal) align.Horizontal = Horizontal;
            if (HasVertical) align.Vertical = Vertical;
            if (TextRotation != null) align.TextRotation = TextRotationToOpenXmlValue(TextRotation.Value);
            if (WrapText != null) align.WrapText = WrapText.Value;
            if (Indent != null) align.Indent = Indent.Value;
            if (RelativeIndent != null) align.RelativeIndent = RelativeIndent.Value;
            if (JustifyLastLine != null) align.JustifyLastLine = JustifyLastLine.Value;
            if (ShrinkToFit != null) align.ShrinkToFit = ShrinkToFit.Value;
            if (HasReadingOrder) align.ReadingOrder = (uint) ReadingOrder;

            return align;
        }

        internal void FromHash(string Hash)
        {
            SetAllNull();

            var sa = Hash.Split(new[] {SLConstants.XmlAlignmentAttributeSeparator}, StringSplitOptions.None);

            if (sa.Length >= 9)
            {
                if (!sa[0].Equals("null"))
                    Horizontal = (HorizontalAlignmentValues) Enum.Parse(typeof(HorizontalAlignmentValues), sa[0]);

                if (!sa[1].Equals("null"))
                    Vertical = (VerticalAlignmentValues) Enum.Parse(typeof(VerticalAlignmentValues), sa[1]);

                if (!sa[2].Equals("null")) TextRotation = int.Parse(sa[2]);

                if (!sa[3].Equals("null")) WrapText = bool.Parse(sa[3]);

                if (!sa[4].Equals("null")) Indent = uint.Parse(sa[4]);

                if (!sa[5].Equals("null")) RelativeIndent = int.Parse(sa[5]);

                if (!sa[6].Equals("null")) JustifyLastLine = bool.Parse(sa[6]);

                if (!sa[7].Equals("null")) ShrinkToFit = bool.Parse(sa[7]);

                if (!sa[8].Equals("null"))
                    ReadingOrder =
                        (SLAlignmentReadingOrderValues) Enum.Parse(typeof(SLAlignmentReadingOrderValues), sa[8]);
            }
        }

        internal string ToHash()
        {
            var sb = new StringBuilder();

            if (HasHorizontal) sb.AppendFormat("{0}{1}", Horizontal, SLConstants.XmlAlignmentAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlAlignmentAttributeSeparator);

            if (HasVertical) sb.AppendFormat("{0}{1}", Vertical, SLConstants.XmlAlignmentAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlAlignmentAttributeSeparator);

            if (TextRotation != null)
                sb.AppendFormat("{0}{1}", TextRotation.Value.ToString(CultureInfo.InvariantCulture),
                    SLConstants.XmlAlignmentAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlAlignmentAttributeSeparator);

            if (WrapText != null) sb.AppendFormat("{0}{1}", WrapText.Value, SLConstants.XmlAlignmentAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlAlignmentAttributeSeparator);

            if (Indent != null)
                sb.AppendFormat("{0}{1}", Indent.Value.ToString(CultureInfo.InvariantCulture),
                    SLConstants.XmlAlignmentAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlAlignmentAttributeSeparator);

            if (RelativeIndent != null)
                sb.AppendFormat("{0}{1}", RelativeIndent.Value.ToString(CultureInfo.InvariantCulture),
                    SLConstants.XmlAlignmentAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlAlignmentAttributeSeparator);

            if (JustifyLastLine != null)
                sb.AppendFormat("{0}{1}", JustifyLastLine.Value, SLConstants.XmlAlignmentAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlAlignmentAttributeSeparator);

            if (ShrinkToFit != null)
                sb.AppendFormat("{0}{1}", ShrinkToFit.Value, SLConstants.XmlAlignmentAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlAlignmentAttributeSeparator);

            if (HasReadingOrder) sb.AppendFormat("{0}{1}", ReadingOrder, SLConstants.XmlAlignmentAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlAlignmentAttributeSeparator);

            return sb.ToString();
        }

        internal int TextRotationToIntuitiveValue(uint Degree)
        {
            var iDegree = 0;

            if ((Degree >= 0) && (Degree <= 90))
                iDegree = (int) Degree;
            else if ((Degree >= 91) && (Degree <= 180))
                iDegree = 90 - (int) Degree;

            return iDegree;
        }

        internal uint TextRotationToOpenXmlValue(int Degree)
        {
            uint iDegree = 0;

            if ((Degree >= 0) && (Degree <= 90))
                iDegree = (uint) Degree;
            else if ((Degree >= -90) && (Degree < 0))
                iDegree = (uint) (90 - Degree);

            return iDegree;
        }

        internal string WriteToXmlTag()
        {
            var sb = new StringBuilder();
            sb.Append("<x:alignment");

            if (HasHorizontal)
                switch (Horizontal)
                {
                    case HorizontalAlignmentValues.Center:
                        sb.Append(" horizontal=\"center\"");
                        break;
                    case HorizontalAlignmentValues.CenterContinuous:
                        sb.Append(" horizontal=\"centerContinuous\"");
                        break;
                    case HorizontalAlignmentValues.Distributed:
                        sb.Append(" horizontal=\"distributed\"");
                        break;
                    case HorizontalAlignmentValues.Fill:
                        sb.Append(" horizontal=\"fill\"");
                        break;
                    case HorizontalAlignmentValues.General:
                        sb.Append(" horizontal=\"general\"");
                        break;
                    case HorizontalAlignmentValues.Justify:
                        sb.Append(" horizontal=\"justify\"");
                        break;
                    case HorizontalAlignmentValues.Left:
                        sb.Append(" horizontal=\"left\"");
                        break;
                    case HorizontalAlignmentValues.Right:
                        sb.Append(" horizontal=\"right\"");
                        break;
                }

            if (HasVertical)
                switch (Vertical)
                {
                    case VerticalAlignmentValues.Bottom:
                        sb.Append(" vertical=\"bottom\"");
                        break;
                    case VerticalAlignmentValues.Center:
                        sb.Append(" vertical=\"center\"");
                        break;
                    case VerticalAlignmentValues.Distributed:
                        sb.Append(" vertical=\"distributed\"");
                        break;
                    case VerticalAlignmentValues.Justify:
                        sb.Append(" vertical=\"justify\"");
                        break;
                    case VerticalAlignmentValues.Top:
                        sb.Append(" vertical=\"top\"");
                        break;
                }

            if (TextRotation != null)
                sb.AppendFormat(" textRotation=\"{0}\"", TextRotation.Value.ToString(CultureInfo.InvariantCulture));
            if (WrapText != null) sb.AppendFormat(" wrapText=\"{0}\"", WrapText.Value ? "1" : "0");
            if (Indent != null) sb.AppendFormat(" indent=\"{0}\"", Indent.Value.ToString(CultureInfo.InvariantCulture));
            if (RelativeIndent != null)
                sb.AppendFormat(" relativeIndent=\"{0}\"", RelativeIndent.Value.ToString(CultureInfo.InvariantCulture));
            if (JustifyLastLine != null) sb.AppendFormat(" justifyLastLine=\"{0}\"", JustifyLastLine.Value ? "1" : "0");
            if (ShrinkToFit != null) sb.AppendFormat(" shrinkToFit=\"{0}\"", ShrinkToFit.Value ? "1" : "0");
            if (HasReadingOrder) sb.AppendFormat(" readingOrder=\"{0}\"", (uint) ReadingOrder);

            sb.Append(" />");

            return sb.ToString();
        }

        internal SLAlignment Clone()
        {
            var align = new SLAlignment();
            align.HasHorizontal = HasHorizontal;
            align.vHorizontal = vHorizontal;
            align.HasVertical = HasVertical;
            align.vVertical = vVertical;
            align.iTextRotation = iTextRotation;
            align.WrapText = WrapText;
            align.Indent = Indent;
            align.RelativeIndent = RelativeIndent;
            align.JustifyLastLine = JustifyLastLine;
            align.ShrinkToFit = ShrinkToFit;
            align.HasReadingOrder = HasReadingOrder;
            align.vReadingOrder = vReadingOrder;

            return align;
        }
    }
}