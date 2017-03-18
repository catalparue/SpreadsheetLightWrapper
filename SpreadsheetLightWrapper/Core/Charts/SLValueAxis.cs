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
    ///     Encapsulates properties and methods for setting value axes in charts.
    ///     This simulates the DocumentFormat.OpenXml.Drawing.Charts.ValueAxis class.
    /// </summary>
    public class SLValueAxis : EGAxShared
    {
        internal SLValueAxis(List<Color> ThemeColors, bool IsStylish = false) : base(ThemeColors, IsStylish)
        {
            CrossBetween = C.CrossBetweenValues.Between;
            MajorUnit = null;
            MinorUnit = null;
            BuiltInUnitValues = null;
            ShowDisplayUnitsLabel = false;

            if (IsStylish)
            {
                ShapeProperties.Fill.SetNoFill();
                ShapeProperties.Outline.SetNoLine();
            }
        }

        // the actual value is stored at the category/date/value axis
        internal C.CrossBetweenValues CrossBetween { get; set; }

        /// <summary>
        ///     The major unit on the axis. A null value means it's automatically set.
        /// </summary>
        public double? MajorUnit { get; set; }

        /// <summary>
        ///     The minor unit on the axis. A null value means it's automatically set.
        /// </summary>
        public double? MinorUnit { get; set; }

        /// <summary>
        ///     Logarithmic scale of the axis, ranging from 2 to 1000 (both inclusive). A null value means it's not used.
        /// </summary>
        public double? LogarithmicScale
        {
            get { return LogBase; }
            set { LogBase = value; }
        }

        // C.DisplayUnits
        internal C.BuiltInUnitValues? BuiltInUnitValues { get; set; }
        internal bool ShowDisplayUnitsLabel { get; set; }

        /// <summary>
        ///     The maximum value on the axis. A null value means it's automatically set.
        /// </summary>
        public double? Maximum
        {
            get { return MaxAxisValue; }
            set { MaxAxisValue = value; }
        }

        /// <summary>
        ///     The minimum value on the axis. A null value means it's automatically set.
        /// </summary>
        public double? Minimum
        {
            get { return MinAxisValue; }
            set { MinAxisValue = value; }
        }

        /// <summary>
        ///     Clear all styling shape properties. Use this if you want to start styling from a clean slate.
        /// </summary>
        public void ClearShapeProperties()
        {
            ShapeProperties = new SLShapeProperties(ShapeProperties.listThemeColors);
        }

        /// <summary>
        ///     Set the display units on the axis.
        /// </summary>
        /// <param name="BuiltInUnit">Built-in unit types.</param>
        /// <param name="ShowDisplayUnitsLabel">True to show the display units label on the chart. False otherwise.</param>
        public void SetDisplayUnits(C.BuiltInUnitValues BuiltInUnit, bool ShowDisplayUnitsLabel)
        {
            BuiltInUnitValues = BuiltInUnit;
            this.ShowDisplayUnitsLabel = ShowDisplayUnitsLabel;
        }

        /// <summary>
        ///     Remove the display units on the axis.
        /// </summary>
        public void RemoveDisplayUnits()
        {
            BuiltInUnitValues = null;
            ShowDisplayUnitsLabel = false;
        }

        /// <summary>
        ///     Set the corresponding category/date/value axis to cross this axis at an automatic value.
        /// </summary>
        public void SetAutomaticOtherAxisCrossing()
        {
            OtherAxisIsCrosses = true;
            OtherAxisCrosses = C.CrossesValues.AutoZero;
            OtherAxisCrossesAt = 0;
        }

        /// <summary>
        ///     Set the corresponding category/date/value axis to cross this axis at a given value.
        /// </summary>
        /// <param name="CrossingAxisValue">Axis value to cross at.</param>
        public void SetOtherAxisCrossing(double CrossingAxisValue)
        {
            OtherAxisIsCrosses = false;
            OtherAxisCrosses = C.CrossesValues.AutoZero;
            OtherAxisCrossesAt = CrossingAxisValue;
        }

        /// <summary>
        ///     Set the corresponding category/date/value axis to cross this axis at the maximum value.
        /// </summary>
        public void SetMaximumOtherAxisCrossing()
        {
            OtherAxisIsCrosses = true;
            OtherAxisCrosses = C.CrossesValues.Maximum;
            OtherAxisCrossesAt = 0;
        }

        internal C.ValueAxis ToValueAxis(bool IsStylish = false)
        {
            var va = new C.ValueAxis();
            va.AxisId = new C.AxisId {Val = AxisId};

            va.Scaling = new C.Scaling();
            va.Scaling.Orientation = new C.Orientation {Val = Orientation};
            if (LogBase != null) va.Scaling.LogBase = new C.LogBase {Val = LogBase.Value};
            if (MaxAxisValue != null) va.Scaling.MaxAxisValue = new C.MaxAxisValue {Val = MaxAxisValue.Value};
            if (MinAxisValue != null) va.Scaling.MinAxisValue = new C.MinAxisValue {Val = MinAxisValue.Value};

            va.Delete = new C.Delete {Val = Delete};

            var axpos = AxisPosition;
            if (!ForceAxisPosition)
            {
                if (OtherAxisIsInReverseOrder) axpos = SLChartTool.GetOppositePosition(axpos);
                if (OtherAxisCrossedAtMaximum) axpos = SLChartTool.GetOppositePosition(axpos);
            }
            va.AxisPosition = new C.AxisPosition {Val = axpos};

            if (ShowMajorGridlines)
                va.MajorGridlines = MajorGridlines.ToMajorGridlines(IsStylish);

            if (ShowMinorGridlines)
                va.MinorGridlines = MinorGridlines.ToMinorGridlines(IsStylish);

            if (ShowTitle)
                va.Title = Title.ToTitle(IsStylish);

            if (HasNumberingFormat)
                va.NumberingFormat = new C.NumberingFormat
                {
                    FormatCode = FormatCode,
                    SourceLinked = SourceLinked
                };

            va.MajorTickMark = new C.MajorTickMark {Val = MajorTickMark};
            va.MinorTickMark = new C.MinorTickMark {Val = MinorTickMark};
            va.TickLabelPosition = new C.TickLabelPosition {Val = TickLabelPosition};

            if (ShapeProperties.HasShapeProperties)
                va.ChartShapeProperties = ShapeProperties.ToChartShapeProperties(IsStylish);

            if ((Rotation != null) || (Vertical != null) || (Anchor != null) || (AnchorCenter != null))
            {
                va.TextProperties = new C.TextProperties();
                va.TextProperties.BodyProperties = new A.BodyProperties();
                if (Rotation != null)
                    va.TextProperties.BodyProperties.Rotation =
                        (int) (Rotation.Value*SLConstants.DegreeToAngleRepresentation);
                if (Vertical != null) va.TextProperties.BodyProperties.Vertical = Vertical.Value;
                if (Anchor != null) va.TextProperties.BodyProperties.Anchor = Anchor.Value;
                if (AnchorCenter != null) va.TextProperties.BodyProperties.AnchorCenter = AnchorCenter.Value;

                va.TextProperties.ListStyle = new A.ListStyle();

                var para = new A.Paragraph();
                para.ParagraphProperties = new A.ParagraphProperties();
                para.ParagraphProperties.Append(new A.DefaultRunProperties());
                va.TextProperties.Append(para);
            }
            else if (IsStylish)
            {
                va.TextProperties = new C.TextProperties();
                va.TextProperties.BodyProperties = new A.BodyProperties
                {
                    Rotation = -60000000,
                    UseParagraphSpacing = true,
                    VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis,
                    Vertical = A.TextVerticalValues.Horizontal,
                    Wrap = A.TextWrappingValues.Square,
                    Anchor = A.TextAnchoringTypeValues.Center,
                    AnchorCenter = true
                };
                va.TextProperties.ListStyle = new A.ListStyle();

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

                va.TextProperties.Append(para);
            }

            va.CrossingAxis = new C.CrossingAxis {Val = CrossingAxis};

            if (IsCrosses != null)
                if (IsCrosses.Value)
                    va.Append(new C.Crosses {Val = Crosses});
                else
                    va.Append(new C.CrossesAt {Val = CrossesAt});

            va.Append(new C.CrossBetween {Val = CrossBetween});
            if (MajorUnit != null) va.Append(new C.MajorUnit {Val = MajorUnit.Value});
            if (MinorUnit != null) va.Append(new C.MinorUnit {Val = MinorUnit.Value});

            if (BuiltInUnitValues != null)
            {
                var du = new C.DisplayUnits();
                du.Append(new C.BuiltInUnit {Val = BuiltInUnitValues.Value});
                if (ShowDisplayUnitsLabel)
                {
                    var dul = new C.DisplayUnitsLabel();
                    dul.Layout = new C.Layout();
                    du.Append(dul);
                }
                va.Append(du);
            }

            return va;
        }

        internal SLValueAxis Clone()
        {
            var va = new SLValueAxis(ShapeProperties.listThemeColors);
            va.Rotation = Rotation;
            va.Vertical = Vertical;
            va.Anchor = Anchor;
            va.AnchorCenter = AnchorCenter;
            va.AxisId = AxisId;
            va.fLogBase = fLogBase;
            va.Orientation = Orientation;
            va.MaxAxisValue = MaxAxisValue;
            va.MinAxisValue = MinAxisValue;
            va.OtherAxisIsInReverseOrder = OtherAxisIsInReverseOrder;
            va.OtherAxisCrossedAtMaximum = OtherAxisCrossedAtMaximum;
            va.Delete = Delete;
            va.ForceAxisPosition = ForceAxisPosition;
            va.AxisPosition = AxisPosition;
            va.ShowMajorGridlines = ShowMajorGridlines;
            va.MajorGridlines = MajorGridlines.Clone();
            va.ShowMinorGridlines = ShowMinorGridlines;
            va.MinorGridlines = MinorGridlines.Clone();
            va.ShowTitle = ShowTitle;
            va.Title = Title.Clone();
            va.HasNumberingFormat = HasNumberingFormat;
            va.sFormatCode = sFormatCode;
            va.bSourceLinked = bSourceLinked;
            va.MajorTickMark = MajorTickMark;
            va.MinorTickMark = MinorTickMark;
            va.TickLabelPosition = TickLabelPosition;
            va.ShapeProperties = ShapeProperties.Clone();
            va.CrossingAxis = CrossingAxis;
            va.IsCrosses = IsCrosses;
            va.Crosses = Crosses;
            va.CrossesAt = CrossesAt;
            va.OtherAxisIsCrosses = OtherAxisIsCrosses;
            va.OtherAxisCrosses = OtherAxisCrosses;
            va.OtherAxisCrossesAt = OtherAxisCrossesAt;

            va.CrossBetween = CrossBetween;
            va.MajorUnit = MajorUnit;
            va.MinorUnit = MinorUnit;
            va.BuiltInUnitValues = BuiltInUnitValues;
            va.ShowDisplayUnitsLabel = ShowDisplayUnitsLabel;

            return va;
        }
    }
}