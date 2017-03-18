using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using SpreadsheetLightWrapper.Core.Drawing;
using SpreadsheetLightWrapper.Core.misc;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLightWrapper.Core.Charts
{
    /// <summary>
    ///     Encapsulates properties and methods for setting series axes in charts.
    ///     This simulates the DocumentFormat.OpenXml.Drawing.Charts.SeriesAxis class.
    /// </summary>
    public class SLSeriesAxis : EGAxShared
    {
        internal ushort iTickLabelSkip;

        internal ushort iTickMarkSkip;

        internal SLSeriesAxis(List<Color> ThemeColors, bool IsStylish = false) : base(ThemeColors, IsStylish)
        {
            iTickLabelSkip = 1;
            iTickMarkSkip = 1;

            if (IsStylish)
            {
                ShapeProperties.Fill.SetNoFill();
                ShapeProperties.Outline.Width = 0.75m;
                ShapeProperties.Outline.CapType = A.LineCapValues.Flat;
                ShapeProperties.Outline.CompoundLineType = A.CompoundLineValues.Single;
                ShapeProperties.Outline.Alignment = A.PenAlignmentValues.Center;
                ShapeProperties.Outline.SetSolidLine(A.SchemeColorValues.Text1, 0.85m, 0);
                ShapeProperties.Outline.JoinType = SLLineJoinValues.Round;
            }
        }

        /// <summary>
        ///     This is the interval between labels, and is at least 1. A suggested range is 1 to 255 (both inclusive).
        /// </summary>
        public ushort TickLabelSkip
        {
            get { return iTickLabelSkip; }
            set
            {
                iTickLabelSkip = value;
                if (iTickLabelSkip < 1) iTickLabelSkip = 1;
            }
        }

        /// <summary>
        ///     This is the interval between tick marks, and is at least 1. A suggested range is 1 to 31999 (both inclusive).
        /// </summary>
        public ushort TickMarkSkip
        {
            get { return iTickMarkSkip; }
            set
            {
                iTickMarkSkip = value;
                if (iTickMarkSkip < 1) iTickMarkSkip = 1;
            }
        }

        internal C.SeriesAxis ToSeriesAxis(bool IsStylish = false)
        {
            var sa = new C.SeriesAxis();
            sa.AxisId = new C.AxisId {Val = AxisId};

            sa.Scaling = new C.Scaling();
            sa.Scaling.Orientation = new C.Orientation {Val = Orientation};
            if (LogBase != null) sa.Scaling.LogBase = new C.LogBase {Val = LogBase.Value};
            if (MaxAxisValue != null) sa.Scaling.MaxAxisValue = new C.MaxAxisValue {Val = MaxAxisValue.Value};
            if (MinAxisValue != null) sa.Scaling.MinAxisValue = new C.MinAxisValue {Val = MinAxisValue.Value};

            sa.Delete = new C.Delete {Val = Delete};

            sa.AxisPosition = new C.AxisPosition {Val = AxisPosition};

            if (ShowMajorGridlines)
                sa.MajorGridlines = MajorGridlines.ToMajorGridlines(IsStylish);

            if (ShowMinorGridlines)
                sa.MinorGridlines = MinorGridlines.ToMinorGridlines(IsStylish);

            if (ShowTitle)
                sa.Title = Title.ToTitle(IsStylish);

            if (HasNumberingFormat)
                sa.NumberingFormat = new C.NumberingFormat
                {
                    FormatCode = FormatCode,
                    SourceLinked = SourceLinked
                };

            sa.MajorTickMark = new C.MajorTickMark {Val = MajorTickMark};
            sa.MinorTickMark = new C.MinorTickMark {Val = MinorTickMark};
            sa.TickLabelPosition = new C.TickLabelPosition {Val = TickLabelPosition};

            if (ShapeProperties.HasShapeProperties) sa.ChartShapeProperties = ShapeProperties.ToChartShapeProperties();

            if ((Rotation != null) || (Vertical != null) || (Anchor != null) || (AnchorCenter != null))
            {
                sa.TextProperties = new C.TextProperties();
                sa.TextProperties.BodyProperties = new A.BodyProperties();
                if (Rotation != null)
                    sa.TextProperties.BodyProperties.Rotation =
                        (int) (Rotation.Value*SLConstants.DegreeToAngleRepresentation);
                if (Vertical != null) sa.TextProperties.BodyProperties.Vertical = Vertical.Value;
                if (Anchor != null) sa.TextProperties.BodyProperties.Anchor = Anchor.Value;
                if (AnchorCenter != null) sa.TextProperties.BodyProperties.AnchorCenter = AnchorCenter.Value;

                sa.TextProperties.ListStyle = new A.ListStyle();

                var para = new A.Paragraph();
                para.ParagraphProperties = new A.ParagraphProperties();
                para.ParagraphProperties.Append(new A.DefaultRunProperties());
                sa.TextProperties.Append(para);
            }
            else if (IsStylish)
            {
                sa.TextProperties = new C.TextProperties();
                sa.TextProperties.BodyProperties = new A.BodyProperties
                {
                    Rotation = -60000000,
                    UseParagraphSpacing = true,
                    VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis,
                    Vertical = A.TextVerticalValues.Horizontal,
                    Wrap = A.TextWrappingValues.Square,
                    Anchor = A.TextAnchoringTypeValues.Center,
                    AnchorCenter = true
                };
                sa.TextProperties.ListStyle = new A.ListStyle();

                var para = new A.Paragraph();
                para.ParagraphProperties = new A.ParagraphProperties();

                var defrunprops = new A.DefaultRunProperties();
                defrunprops.FontSize = 900;
                defrunprops.Bold = false;
                defrunprops.Italic = false;
                defrunprops.Underline = A.TextUnderlineValues.None;
                defrunprops.Strike = A.TextStrikeValues.NoStrike;
                defrunprops.Kerning = 1200;
                defrunprops.Baseline = 0;

                var schclr = new A.SchemeColor {Val = A.SchemeColorValues.Text1};
                schclr.Append(new A.LuminanceModulation {Val = 65000});
                schclr.Append(new A.LuminanceOffset {Val = 35000});
                defrunprops.Append(new A.SolidFill
                {
                    SchemeColor = schclr
                });

                defrunprops.Append(new A.LatinFont {Typeface = "+mn-lt"});
                defrunprops.Append(new A.EastAsianFont {Typeface = "+mn-ea"});
                defrunprops.Append(new A.ComplexScriptFont {Typeface = "+mn-cs"});

                para.ParagraphProperties.Append(defrunprops);
                para.Append(new A.EndParagraphRunProperties {Language = CultureInfo.CurrentCulture.Name});

                sa.TextProperties.Append(para);
            }

            sa.CrossingAxis = new C.CrossingAxis {Val = CrossingAxis};

            if (IsCrosses != null)
                if (IsCrosses.Value)
                    sa.Append(new C.Crosses {Val = Crosses});
                else
                    sa.Append(new C.CrossesAt {Val = CrossesAt});

            if (iTickLabelSkip > 1) sa.Append(new C.TickLabelSkip {Val = TickLabelSkip});
            if (iTickMarkSkip > 1) sa.Append(new C.TickMarkSkip {Val = TickMarkSkip});

            return sa;
        }

        internal SLSeriesAxis Clone()
        {
            var sa = new SLSeriesAxis(ShapeProperties.listThemeColors);
            sa.Rotation = Rotation;
            sa.Vertical = Vertical;
            sa.Anchor = Anchor;
            sa.AnchorCenter = AnchorCenter;
            sa.AxisId = AxisId;
            sa.fLogBase = fLogBase;
            sa.Orientation = Orientation;
            sa.MaxAxisValue = MaxAxisValue;
            sa.MinAxisValue = MinAxisValue;
            sa.OtherAxisIsInReverseOrder = OtherAxisIsInReverseOrder;
            sa.OtherAxisCrossedAtMaximum = OtherAxisCrossedAtMaximum;
            sa.Delete = Delete;
            sa.ForceAxisPosition = ForceAxisPosition;
            sa.AxisPosition = AxisPosition;
            sa.ShowMajorGridlines = ShowMajorGridlines;
            sa.MajorGridlines = MajorGridlines.Clone();
            sa.ShowMinorGridlines = ShowMinorGridlines;
            sa.MinorGridlines = MinorGridlines.Clone();
            sa.ShowTitle = ShowTitle;
            sa.Title = Title.Clone();
            sa.HasNumberingFormat = HasNumberingFormat;
            sa.sFormatCode = sFormatCode;
            sa.bSourceLinked = bSourceLinked;
            sa.MajorTickMark = MajorTickMark;
            sa.MinorTickMark = MinorTickMark;
            sa.TickLabelPosition = TickLabelPosition;
            sa.ShapeProperties = ShapeProperties.Clone();
            sa.CrossingAxis = CrossingAxis;
            sa.IsCrosses = IsCrosses;
            sa.Crosses = Crosses;
            sa.CrossesAt = CrossesAt;
            sa.OtherAxisIsCrosses = OtherAxisIsCrosses;
            sa.OtherAxisCrosses = OtherAxisCrosses;
            sa.OtherAxisCrossesAt = OtherAxisCrossesAt;

            sa.iTickLabelSkip = iTickLabelSkip;
            sa.iTickMarkSkip = iTickMarkSkip;

            return sa;
        }
    }
}