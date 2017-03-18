using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using SpreadsheetLightWrapper.Core.Drawing;
using SpreadsheetLightWrapper.Core.misc;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLightWrapper.Core.Charts
{
    internal enum SLAxisType
    {
        Category,
        Date,
        Value
    }

    /// <summary>
    ///     Encapsulates properties and methods for setting chart axes, specifically simulating
    ///     DocumentFormat.OpenXml.Drawing.Charts.CategoryAxis,
    ///     DocumentFormat.OpenXml.Drawing.Charts.DateAxis and
    ///     DocumentFormat.OpenXml.Drawing.Charts.ValueAxis classes.
    /// </summary>
    public class SLTextAxis : EGAxShared
    {
        internal ushort iLabelOffset;

        internal int? iMajorUnit;

        internal int? iMinorUnit;

        internal ushort iTickLabelSkip;

        internal ushort iTickMarkSkip;
        internal C.TimeUnitValues vMajorTimeUnit;
        internal C.TimeUnitValues vMinorTimeUnit;

        internal SLTextAxis(List<Color> ThemeColors, bool Date1904, bool IsStylish = false)
            : base(ThemeColors, IsStylish)
        {
            this.Date1904 = Date1904;

            AxisType = SLAxisType.Category;
            AutoLabeled = true;

            iTickLabelSkip = 1;
            iTickMarkSkip = 1;
            iLabelOffset = 100;

            ValueMajorUnit = null;
            ValueMinorUnit = null;
            BuiltInUnitValues = null;
            ShowDisplayUnitsLabel = false;

            BaseUnit = null;
            iMajorUnit = null;
            vMajorTimeUnit = C.TimeUnitValues.Days;
            iMinorUnit = null;
            vMinorTimeUnit = C.TimeUnitValues.Days;

            CrossBetween = C.CrossBetweenValues.Between;
            LabelAlignment = C.LabelAlignmentValues.Center;

            // it used to be true. I have no idea what this does...
            NoMultiLevelLabels = false;

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

        internal bool Date1904 { get; set; }

        internal SLAxisType AxisType { get; set; }

        // switch when axis types are changed (category to date or date to category)
        internal bool AutoLabeled { get; set; }

        /// <summary>
        ///     This is the interval between labels, and is at least 1. A suggested range is 1 to 255 (both inclusive). This is
        ///     only for category axes.
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
        ///     This is the interval between tick marks, and is at least 1. A suggested range is 1 to 31999 (both inclusive). This
        ///     is only for category axes.
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

        /// <summary>
        ///     Label alignment for the category axis. This is ignored for date axes.
        /// </summary>
        public C.LabelAlignmentValues LabelAlignment { get; set; }

        /// <summary>
        ///     This is the label distance from the axis, ranging from 0 to 1000 (both inclusive). The default is 100.
        /// </summary>
        public ushort LabelOffset
        {
            get { return iLabelOffset; }
            set
            {
                iLabelOffset = value;
                if (iLabelOffset > 1000) iLabelOffset = 1000;
            }
        }

        /// <summary>
        ///     The maximum value on the axis. A null value means it's automatically set. WARNING: This is used for date axes. It's
        ///     also shared with value axes. If it's set for category axes, chart behaviour is not defined.
        /// </summary>
        public DateTime? MaximumDate
        {
            get
            {
                if (MaxAxisValue == null)
                    return null;
                return SLTool.CalculateDateTimeFromDaysFromEpoch(MaxAxisValue.Value, Date1904);
            }
            set
            {
                if (value == null)
                    MaxAxisValue = null;
                else
                    MaxAxisValue = SLTool.CalculateDaysFromEpoch(value.Value, Date1904);
            }
        }

        /// <summary>
        ///     The minimum value on the axis. A null value means it's automatically set. WARNING: This is used for date axes. It's
        ///     also shared with value axes. If it's set for category axes, chart behaviour is not defined.
        /// </summary>
        public DateTime? MinimumDate
        {
            get
            {
                if (MinAxisValue == null)
                    return null;
                return SLTool.CalculateDateTimeFromDaysFromEpoch(MinAxisValue.Value, Date1904);
            }
            set
            {
                if (value == null)
                    MinAxisValue = null;
                else
                    MinAxisValue = SLTool.CalculateDaysFromEpoch(value.Value, Date1904);
            }
        }

        /// <summary>
        ///     The maximum value on the axis. A null value means it's automatically set. WARNING: This is used for value axis.
        ///     It's also shared with date axes. If it's set for category axes, chart behaviour is not defined.
        /// </summary>
        public double? MaximumValue
        {
            get { return MaxAxisValue; }
            set { MaxAxisValue = value; }
        }

        /// <summary>
        ///     The minimum value on the axis. A null value means it's automatically set. WARNING: This is used for value axis.
        ///     It's also shared with date axes. If it's set for category axes, chart behaviour is not defined.
        /// </summary>
        public double? MinimumValue
        {
            get { return MinAxisValue; }
            set { MinAxisValue = value; }
        }

        /// <summary>
        ///     The major unit on the axis. A null value means it's automatically set. This is for the value axis.
        /// </summary>
        public double? ValueMajorUnit { get; set; }

        /// <summary>
        ///     The minor unit on the axis. A null value means it's automatically set. This is for the value axis.
        /// </summary>
        public double? ValueMinorUnit { get; set; }

        /// <summary>
        ///     Logarithmic scale of the axis, ranging from 2 to 1000 (both inclusive). A null value means it's not used. This is
        ///     for the value axis.
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
        ///     The base unit for date axes. A null value means it's automatically set.
        /// </summary>
        public C.TimeUnitValues? BaseUnit { get; set; }

        // This is actually for the value axis, but due to the way Excel displays the user interface,
        // this is set on the category/date/value axis settings. I don't understand it either...
        /// <summary>
        ///     This sets how the axis crosses regarding the tick marks (or position of the axis). Use Between for "between tick
        ///     marks", and MidpointCategory for "on tick marks".
        /// </summary>
        public C.CrossBetweenValues CrossBetween { get; set; }

        /// <summary>
        ///     Indicates if labels are shown as flat text. If false, then the labels are shown as a hierarchy.
        ///     This is used only for category axes. The default is true.
        /// </summary>
        public bool NoMultiLevelLabels { get; set; }

        /// <summary>
        ///     Clear all styling shape properties. Use this if you want to start styling from a clean slate.
        /// </summary>
        public void ClearShapeProperties()
        {
            ShapeProperties = new SLShapeProperties(ShapeProperties.listThemeColors);
        }

        // We have SetAsCategoryAxis() and SetAsDateAxis() because
        // we want to keep the option of SetAutomaticAxisType()
        // and that needs examining the chart data to determine the type.
        // I don't feel that's very value-added enough...

        /// <summary>
        ///     Set this axis as a category axis. WARNING: This only works if it's a category/date axis. This fails if it's already
        ///     a value axis.
        /// </summary>
        public void SetAsCategoryAxis()
        {
            if (AxisType != SLAxisType.Value)
            {
                AxisType = SLAxisType.Category;
                AutoLabeled = false;
            }
        }

        /// <summary>
        ///     Set this axis as a date axis. WARNING: This only works if it's a category/date axis. This fails if it's already a
        ///     value axis.
        /// </summary>
        public void SetAsDateAxis()
        {
            if (AxisType != SLAxisType.Value)
            {
                AxisType = SLAxisType.Date;
                AutoLabeled = false;
            }
        }

        /// <summary>
        ///     Set the major unit for date axes to be automatic.
        /// </summary>
        public void SetAutomaticDateMajorUnit()
        {
            iMajorUnit = null;
            vMajorTimeUnit = C.TimeUnitValues.Days;
        }

        /// <summary>
        ///     Set the major unit for date axes.
        /// </summary>
        /// <param name="MajorUnit">A positive value. Suggested range is 1 to 999999999 (both inclusive).</param>
        /// <param name="MajorTimeUnit">The time unit.</param>
        public void SetDateMajorUnit(int MajorUnit, C.TimeUnitValues MajorTimeUnit)
        {
            iMajorUnit = MajorUnit;
            vMajorTimeUnit = MajorTimeUnit;
        }

        /// <summary>
        ///     Set the minor unit for date axes to be automatic.
        /// </summary>
        public void SetAutomaticDateMinorUnit()
        {
            iMinorUnit = null;
            vMinorTimeUnit = C.TimeUnitValues.Days;
        }

        /// <summary>
        ///     Set the minor unit for date axes.
        /// </summary>
        /// <param name="MinorUnit">A positive value. Suggested range is 1 to 999999999 (both inclusive).</param>
        /// <param name="MinorTimeUnit">The time unit.</param>
        public void SetDateMinorUnit(int MinorUnit, C.TimeUnitValues MinorTimeUnit)
        {
            iMinorUnit = MinorUnit;
            vMinorTimeUnit = MinorTimeUnit;
        }

        /// <summary>
        ///     Set the display units on the axis. This is for value axis.
        /// </summary>
        /// <param name="BuiltInUnit">Built-in unit types.</param>
        /// <param name="ShowDisplayUnitsLabel">True to show the display units label on the chart. False otherwise.</param>
        public void SetDisplayUnits(C.BuiltInUnitValues BuiltInUnit, bool ShowDisplayUnitsLabel)
        {
            BuiltInUnitValues = BuiltInUnit;
            this.ShowDisplayUnitsLabel = ShowDisplayUnitsLabel;
        }

        /// <summary>
        ///     Remove the display units on the axis. This is for value axis.
        /// </summary>
        public void RemoveDisplayUnits()
        {
            BuiltInUnitValues = null;
            ShowDisplayUnitsLabel = false;
        }

        /// <summary>
        ///     Set the corresponding value axis to cross this axis at an automatic value.
        /// </summary>
        public void SetAutomaticOtherAxisCrossing()
        {
            OtherAxisIsCrosses = true;
            OtherAxisCrosses = C.CrossesValues.AutoZero;
            OtherAxisCrossesAt = 0;
        }

        /// <summary>
        ///     Set the corresponding value axis to cross this axis at a given category number. Suggested range is 1 to 31999 (both
        ///     inclusive). This is for category axis. WARNING: Internally, this is used for category, date and value axes.
        ///     Remember to set the axis type.
        /// </summary>
        /// <param name="CategoryNumber">Category number to cross at.</param>
        public void SetOtherAxisCrossing(int CategoryNumber)
        {
            OtherAxisIsCrosses = false;
            OtherAxisCrosses = C.CrossesValues.AutoZero;
            OtherAxisCrossesAt = CategoryNumber;
        }

        /// <summary>
        ///     Set the corresponding value axis to cross this axis at a given date. This is for date axis. WARNING: Internally,
        ///     this is used for category, date and value axes. Remember to set the axis type.
        /// </summary>
        /// <param name="DateToBeCrossed">Date to cross at.</param>
        public void SetOtherAxisCrossing(DateTime DateToBeCrossed)
        {
            OtherAxisIsCrosses = false;
            OtherAxisCrosses = C.CrossesValues.AutoZero;
            OtherAxisCrossesAt = SLTool.CalculateDaysFromEpoch(DateToBeCrossed, Date1904);
            // the given date is before the epochs (1900 or 1904).
            // Just set to whatever the current epoch is being used.
            if (OtherAxisCrossesAt < 0.0) OtherAxisCrossesAt = Date1904 ? 0.0 : 1.0;
        }

        /// <summary>
        ///     Set the corresponding value axis to cross this axis at a given value. This is for value axis. WARNING: Internally,
        ///     this is used for category, date and value axes. If it's already a value axis, you can't set the axis type.
        /// </summary>
        /// <param name="CrossingAxisValue">Axis value to cross at.</param>
        public void SetOtherAxisCrossing(double CrossingAxisValue)
        {
            OtherAxisIsCrosses = false;
            OtherAxisCrosses = C.CrossesValues.AutoZero;
            OtherAxisCrossesAt = CrossingAxisValue;
        }

        /// <summary>
        ///     Set the corresponding value axis to cross this axis at the maximum value.
        /// </summary>
        public void SetMaximumOtherAxisCrossing()
        {
            OtherAxisIsCrosses = true;
            OtherAxisCrosses = C.CrossesValues.Maximum;
            OtherAxisCrossesAt = 0;
        }

        internal C.CategoryAxis ToCategoryAxis(bool IsStylish = false)
        {
            var ca = new C.CategoryAxis();
            ca.AxisId = new C.AxisId {Val = AxisId};

            ca.Scaling = new C.Scaling();
            ca.Scaling.Orientation = new C.Orientation {Val = Orientation};
            if (LogBase != null) ca.Scaling.LogBase = new C.LogBase {Val = LogBase.Value};
            if (MaxAxisValue != null) ca.Scaling.MaxAxisValue = new C.MaxAxisValue {Val = MaxAxisValue.Value};
            if (MinAxisValue != null) ca.Scaling.MinAxisValue = new C.MinAxisValue {Val = MinAxisValue.Value};

            ca.Delete = new C.Delete {Val = Delete};

            var axpos = AxisPosition;
            if (!ForceAxisPosition)
            {
                if (OtherAxisIsInReverseOrder) axpos = SLChartTool.GetOppositePosition(axpos);
                if (OtherAxisCrossedAtMaximum) axpos = SLChartTool.GetOppositePosition(axpos);
            }
            ca.AxisPosition = new C.AxisPosition {Val = axpos};

            if (ShowMajorGridlines)
                ca.MajorGridlines = MajorGridlines.ToMajorGridlines(IsStylish);

            if (ShowMinorGridlines)
                ca.MinorGridlines = MinorGridlines.ToMinorGridlines(IsStylish);

            if (ShowTitle)
                ca.Title = Title.ToTitle(IsStylish);

            if (HasNumberingFormat)
                ca.NumberingFormat = new C.NumberingFormat
                {
                    FormatCode = FormatCode,
                    SourceLinked = SourceLinked
                };

            ca.MajorTickMark = new C.MajorTickMark {Val = MajorTickMark};
            ca.MinorTickMark = new C.MinorTickMark {Val = MinorTickMark};
            ca.TickLabelPosition = new C.TickLabelPosition {Val = TickLabelPosition};

            if (ShapeProperties.HasShapeProperties)
                ca.ChartShapeProperties = ShapeProperties.ToChartShapeProperties(IsStylish);

            if ((Rotation != null) || (Vertical != null) || (Anchor != null) || (AnchorCenter != null))
            {
                ca.TextProperties = new C.TextProperties();
                ca.TextProperties.BodyProperties = new A.BodyProperties();
                if (Rotation != null)
                    ca.TextProperties.BodyProperties.Rotation =
                        (int) (Rotation.Value*SLConstants.DegreeToAngleRepresentation);
                if (Vertical != null) ca.TextProperties.BodyProperties.Vertical = Vertical.Value;
                if (Anchor != null) ca.TextProperties.BodyProperties.Anchor = Anchor.Value;
                if (AnchorCenter != null) ca.TextProperties.BodyProperties.AnchorCenter = AnchorCenter.Value;

                ca.TextProperties.ListStyle = new A.ListStyle();

                var para = new A.Paragraph();
                para.ParagraphProperties = new A.ParagraphProperties();
                para.ParagraphProperties.Append(new A.DefaultRunProperties());
                ca.TextProperties.Append(para);
            }
            else if (IsStylish)
            {
                ca.TextProperties = new C.TextProperties();
                ca.TextProperties.BodyProperties = new A.BodyProperties
                {
                    Rotation = -60000000,
                    UseParagraphSpacing = true,
                    VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis,
                    Vertical = A.TextVerticalValues.Horizontal,
                    Wrap = A.TextWrappingValues.Square,
                    Anchor = A.TextAnchoringTypeValues.Center,
                    AnchorCenter = true
                };
                ca.TextProperties.ListStyle = new A.ListStyle();

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

                ca.TextProperties.Append(para);
            }

            ca.CrossingAxis = new C.CrossingAxis {Val = CrossingAxis};

            if (IsCrosses != null)
                if (IsCrosses.Value)
                    ca.Append(new C.Crosses {Val = Crosses});
                else
                    ca.Append(new C.CrossesAt {Val = CrossesAt});

            ca.Append(new C.AutoLabeled {Val = AutoLabeled});
            ca.Append(new C.LabelAlignment {Val = LabelAlignment});
            ca.Append(new C.LabelOffset {Val = LabelOffset});

            if (iTickLabelSkip > 1) ca.Append(new C.TickLabelSkip {Val = TickLabelSkip});
            if (iTickMarkSkip > 1) ca.Append(new C.TickMarkSkip {Val = TickMarkSkip});

            ca.Append(new C.NoMultiLevelLabels {Val = NoMultiLevelLabels});

            return ca;
        }

        internal C.DateAxis ToDateAxis(bool IsStylish = false)
        {
            var da = new C.DateAxis();
            da.AxisId = new C.AxisId {Val = AxisId};

            da.Scaling = new C.Scaling();
            da.Scaling.Orientation = new C.Orientation {Val = Orientation};
            if (LogBase != null) da.Scaling.LogBase = new C.LogBase {Val = LogBase.Value};
            if (MaxAxisValue != null) da.Scaling.MaxAxisValue = new C.MaxAxisValue {Val = MaxAxisValue.Value};
            if (MinAxisValue != null) da.Scaling.MinAxisValue = new C.MinAxisValue {Val = MinAxisValue.Value};

            da.Delete = new C.Delete {Val = Delete};

            var axpos = AxisPosition;
            if (!ForceAxisPosition)
            {
                if (OtherAxisIsInReverseOrder) axpos = SLChartTool.GetOppositePosition(axpos);
                if (OtherAxisCrossedAtMaximum) axpos = SLChartTool.GetOppositePosition(axpos);
            }
            da.AxisPosition = new C.AxisPosition {Val = axpos};

            if (ShowMajorGridlines)
                da.MajorGridlines = MajorGridlines.ToMajorGridlines(IsStylish);

            if (ShowMinorGridlines)
                da.MinorGridlines = MinorGridlines.ToMinorGridlines(IsStylish);

            if (ShowTitle)
                da.Title = Title.ToTitle(IsStylish);

            if (HasNumberingFormat)
                da.NumberingFormat = new C.NumberingFormat
                {
                    FormatCode = FormatCode,
                    SourceLinked = SourceLinked
                };

            da.MajorTickMark = new C.MajorTickMark {Val = MajorTickMark};
            da.MinorTickMark = new C.MinorTickMark {Val = MinorTickMark};
            da.TickLabelPosition = new C.TickLabelPosition {Val = TickLabelPosition};

            if (ShapeProperties.HasShapeProperties)
                da.ChartShapeProperties = ShapeProperties.ToChartShapeProperties(IsStylish);

            if ((Rotation != null) || (Vertical != null) || (Anchor != null) || (AnchorCenter != null))
            {
                da.TextProperties = new C.TextProperties();
                da.TextProperties.BodyProperties = new A.BodyProperties();
                if (Rotation != null)
                    da.TextProperties.BodyProperties.Rotation =
                        (int) (Rotation.Value*SLConstants.DegreeToAngleRepresentation);
                if (Vertical != null) da.TextProperties.BodyProperties.Vertical = Vertical.Value;
                if (Anchor != null) da.TextProperties.BodyProperties.Anchor = Anchor.Value;
                if (AnchorCenter != null) da.TextProperties.BodyProperties.AnchorCenter = AnchorCenter.Value;

                da.TextProperties.ListStyle = new A.ListStyle();

                var para = new A.Paragraph();
                para.ParagraphProperties = new A.ParagraphProperties();
                para.ParagraphProperties.Append(new A.DefaultRunProperties());
                da.TextProperties.Append(para);
            }
            else if (IsStylish)
            {
                da.TextProperties = new C.TextProperties();
                da.TextProperties.BodyProperties = new A.BodyProperties
                {
                    Rotation = -60000000,
                    UseParagraphSpacing = true,
                    VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis,
                    Vertical = A.TextVerticalValues.Horizontal,
                    Wrap = A.TextWrappingValues.Square,
                    Anchor = A.TextAnchoringTypeValues.Center,
                    AnchorCenter = true
                };
                da.TextProperties.ListStyle = new A.ListStyle();

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

                da.TextProperties.Append(para);
            }

            da.CrossingAxis = new C.CrossingAxis {Val = CrossingAxis};

            if (IsCrosses != null)
                if (IsCrosses.Value)
                    da.Append(new C.Crosses {Val = Crosses});
                else
                    da.Append(new C.CrossesAt {Val = CrossesAt});

            da.Append(new C.AutoLabeled {Val = AutoLabeled});
            da.Append(new C.LabelOffset {Val = LabelOffset});

            if (BaseUnit != null) da.Append(new C.BaseTimeUnit {Val = BaseUnit.Value});

            if (iMajorUnit != null)
            {
                da.Append(new C.MajorUnit {Val = iMajorUnit.Value});
                da.Append(new C.MajorTimeUnit {Val = vMajorTimeUnit});
            }

            if (iMinorUnit != null)
            {
                da.Append(new C.MinorUnit {Val = iMinorUnit.Value});
                da.Append(new C.MinorTimeUnit {Val = vMinorTimeUnit});
            }

            return da;
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
            if (ValueMajorUnit != null) va.Append(new C.MajorUnit {Val = ValueMajorUnit.Value});
            if (ValueMinorUnit != null) va.Append(new C.MinorUnit {Val = ValueMinorUnit.Value});

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

        internal SLTextAxis Clone()
        {
            var ta = new SLTextAxis(ShapeProperties.listThemeColors, Date1904);
            ta.Rotation = Rotation;
            ta.Vertical = Vertical;
            ta.Anchor = Anchor;
            ta.AnchorCenter = AnchorCenter;
            ta.AxisId = AxisId;
            ta.fLogBase = fLogBase;
            ta.Orientation = Orientation;
            ta.MaxAxisValue = MaxAxisValue;
            ta.MinAxisValue = MinAxisValue;
            ta.OtherAxisIsInReverseOrder = OtherAxisIsInReverseOrder;
            ta.OtherAxisCrossedAtMaximum = OtherAxisCrossedAtMaximum;
            ta.Delete = Delete;
            ta.ForceAxisPosition = ForceAxisPosition;
            ta.AxisPosition = AxisPosition;
            ta.ShowMajorGridlines = ShowMajorGridlines;
            ta.MajorGridlines = MajorGridlines.Clone();
            ta.ShowMinorGridlines = ShowMinorGridlines;
            ta.MinorGridlines = MinorGridlines.Clone();
            ta.ShowTitle = ShowTitle;
            ta.Title = Title.Clone();
            ta.HasNumberingFormat = HasNumberingFormat;
            ta.sFormatCode = sFormatCode;
            ta.bSourceLinked = bSourceLinked;
            ta.MajorTickMark = MajorTickMark;
            ta.MinorTickMark = MinorTickMark;
            ta.TickLabelPosition = TickLabelPosition;
            ta.ShapeProperties = ShapeProperties.Clone();
            ta.CrossingAxis = CrossingAxis;
            ta.IsCrosses = IsCrosses;
            ta.Crosses = Crosses;
            ta.CrossesAt = CrossesAt;
            ta.OtherAxisIsCrosses = OtherAxisIsCrosses;
            ta.OtherAxisCrosses = OtherAxisCrosses;
            ta.OtherAxisCrossesAt = OtherAxisCrossesAt;

            ta.Date1904 = Date1904;
            ta.AxisType = AxisType;
            ta.AutoLabeled = AutoLabeled;
            ta.iTickLabelSkip = iTickLabelSkip;
            ta.iTickMarkSkip = iTickMarkSkip;
            ta.LabelAlignment = LabelAlignment;
            ta.iLabelOffset = iLabelOffset;
            ta.ValueMajorUnit = ValueMajorUnit;
            ta.ValueMinorUnit = ValueMinorUnit;
            ta.BuiltInUnitValues = BuiltInUnitValues;
            ta.ShowDisplayUnitsLabel = ShowDisplayUnitsLabel;
            ta.BaseUnit = BaseUnit;
            ta.iMajorUnit = iMajorUnit;
            ta.vMajorTimeUnit = vMajorTimeUnit;
            ta.iMinorUnit = iMinorUnit;
            ta.vMinorTimeUnit = vMinorTimeUnit;
            ta.CrossBetween = CrossBetween;
            ta.NoMultiLevelLabels = NoMultiLevelLabels;

            return ta;
        }
    }
}