using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using Ups.Toolkit.SpreadsheetLight.Core.Drawing;
using Ups.Toolkit.SpreadsheetLight.Core.style;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using SLFill = Ups.Toolkit.SpreadsheetLight.Core.Drawing.SLFill;

namespace Ups.Toolkit.SpreadsheetLight.Core.Charts
{
    /// <summary>
    ///     Encapsulates properties and methods for setting the data table of charts.
    /// </summary>
    public class SLDataTable
    {
        internal SLShapeProperties ShapeProperties;

        internal SLDataTable(List<Color> ThemeColors, bool IsStylish = false)
        {
            ShowHorizontalBorder = true;
            ShowVerticalBorder = true;
            ShowOutlineBorder = true;
            ShowLegendKeys = true;
            ShapeProperties = new SLShapeProperties(ThemeColors);

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

            Font = null;
        }

        /// <summary>
        ///     Specifies if horizontal table borders are shown.
        /// </summary>
        public bool ShowHorizontalBorder { get; set; }

        /// <summary>
        ///     Specifies if vertical table borders are shown.
        /// </summary>
        public bool ShowVerticalBorder { get; set; }

        /// <summary>
        ///     Specifies if table outline borders are shown.
        /// </summary>
        public bool ShowOutlineBorder { get; set; }

        /// <summary>
        ///     Specifies if legend keys are shown.
        /// </summary>
        public bool ShowLegendKeys { get; set; }

        /// <summary>
        ///     Fill properties.
        /// </summary>
        public SLFill Fill
        {
            get { return ShapeProperties.Fill; }
        }

        /// <summary>
        ///     Border properties.
        /// </summary>
        public SLLinePropertiesType Border
        {
            get { return ShapeProperties.Outline; }
        }

        /// <summary>
        ///     Shadow properties.
        /// </summary>
        public SLShadowEffect Shadow
        {
            get { return ShapeProperties.EffectList.Shadow; }
        }

        /// <summary>
        ///     Glow properties.
        /// </summary>
        public SLGlow Glow
        {
            get { return ShapeProperties.EffectList.Glow; }
        }

        /// <summary>
        ///     Soft edge properties.
        /// </summary>
        public SLSoftEdge SoftEdge
        {
            get { return ShapeProperties.EffectList.SoftEdge; }
        }

        /// <summary>
        ///     3D format properties.
        /// </summary>
        public SLFormat3D Format3D
        {
            get { return ShapeProperties.Format3D; }
        }

        internal SLFont Font { get; set; }

        /// <summary>
        ///     Clear all styling shape properties. Use this if you want to start styling from a clean slate.
        /// </summary>
        public void ClearShapeProperties()
        {
            ShapeProperties = new SLShapeProperties(ShapeProperties.listThemeColors);
        }

        /// <summary>
        ///     Set font settings for the contents of the data table.
        /// </summary>
        /// <param name="Font">The SLFont containing the font settings.</param>
        public void SetFont(SLFont Font)
        {
            this.Font = Font.Clone();
        }

        internal C.DataTable ToDataTable(bool IsStylish = false)
        {
            var dt = new C.DataTable();

            if (ShowHorizontalBorder) dt.ShowHorizontalBorder = new C.ShowHorizontalBorder {Val = true};
            if (ShowVerticalBorder) dt.ShowVerticalBorder = new C.ShowVerticalBorder {Val = true};
            if (ShowOutlineBorder) dt.ShowOutlineBorder = new C.ShowOutlineBorder {Val = true};
            if (ShowLegendKeys) dt.ShowKeys = new C.ShowKeys {Val = true};

            if (ShapeProperties.HasShapeProperties)
                dt.ChartShapeProperties = ShapeProperties.ToChartShapeProperties(IsStylish);

            if (Font != null)
            {
                dt.TextProperties = new C.TextProperties();
                dt.TextProperties.BodyProperties = new A.BodyProperties();
                dt.TextProperties.ListStyle = new A.ListStyle();

                dt.TextProperties.Append(Font.ToParagraph());
            }
            else if (IsStylish)
            {
                dt.TextProperties = new C.TextProperties();
                dt.TextProperties.BodyProperties = new A.BodyProperties
                {
                    Rotation = 0,
                    UseParagraphSpacing = true,
                    VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis,
                    Vertical = A.TextVerticalValues.Horizontal,
                    Wrap = A.TextWrappingValues.Square,
                    Anchor = A.TextAnchoringTypeValues.Center,
                    AnchorCenter = true
                };
                dt.TextProperties.ListStyle = new A.ListStyle();

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

                dt.TextProperties.Append(para);
            }

            return dt;
        }

        internal SLDataTable Clone()
        {
            var dt = new SLDataTable(ShapeProperties.listThemeColors);
            dt.ShapeProperties = ShapeProperties.Clone();
            dt.ShowHorizontalBorder = ShowHorizontalBorder;
            dt.ShowVerticalBorder = ShowVerticalBorder;
            dt.ShowOutlineBorder = ShowOutlineBorder;
            dt.ShowLegendKeys = ShowLegendKeys;
            if (Font != null) dt.Font = Font.Clone();

            return dt;
        }
    }
}