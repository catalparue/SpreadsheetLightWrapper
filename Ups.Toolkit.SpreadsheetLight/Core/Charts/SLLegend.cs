using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using Ups.Toolkit.SpreadsheetLight.Core.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Ups.Toolkit.SpreadsheetLight.Core.Charts
{
    /// <summary>
    ///     Encapsulates properties and methods for setting chart legends.
    ///     This simulates the DocumentFormat.OpenXml.Drawing.Charts.Legend class.
    /// </summary>
    public class SLLegend
    {
        internal SLShapeProperties ShapeProperties;

        internal SLLegend(List<Color> ThemeColors, bool IsStylish = false)
        {
            LegendPosition = IsStylish ? C.LegendPositionValues.Bottom : C.LegendPositionValues.Right;
            Overlay = false;
            ShapeProperties = new SLShapeProperties(ThemeColors);

            if (IsStylish)
            {
                ShapeProperties.Fill.SetNoFill();
                ShapeProperties.Outline.SetNoLine();
            }
            else
            {
                ShapeProperties.Fill.BlipDpi = 0;
                ShapeProperties.Fill.BlipRotateWithShape = true;
            }
        }

        /// <summary>
        ///     The position of the legend.
        /// </summary>
        public C.LegendPositionValues LegendPosition { get; set; }

        /// <summary>
        ///     Specifies if the legend is overlayed. True if the legend overlaps the plot area, false otherwise.
        /// </summary>
        public bool Overlay { get; set; }

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
        ///     Clear all styling shape properties. Use this if you want to start styling from a clean slate.
        /// </summary>
        public void ClearShapeProperties()
        {
            ShapeProperties = new SLShapeProperties(ShapeProperties.listThemeColors);
        }

        internal C.Legend ToLegend(bool IsStylish = false)
        {
            var l = new C.Legend();
            l.LegendPosition = new C.LegendPosition {Val = LegendPosition};

            l.Append(new C.Layout());
            l.Append(new C.Overlay {Val = Overlay});

            if (ShapeProperties.HasShapeProperties) l.Append(ShapeProperties.ToChartShapeProperties(IsStylish));

            if (IsStylish)
            {
                var tp = new C.TextProperties();
                tp.BodyProperties = new A.BodyProperties
                {
                    Rotation = 0,
                    UseParagraphSpacing = true,
                    VerticalOverflow = A.TextVerticalOverflowValues.Ellipsis,
                    Vertical = A.TextVerticalValues.Horizontal,
                    Wrap = A.TextWrappingValues.Square,
                    Anchor = A.TextAnchoringTypeValues.Center,
                    AnchorCenter = true
                };
                tp.ListStyle = new A.ListStyle();

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

                tp.Append(para);

                l.Append(tp);
            }

            return l;
        }

        internal SLLegend Clone()
        {
            var l = new SLLegend(ShapeProperties.listThemeColors);
            l.LegendPosition = LegendPosition;
            l.Overlay = Overlay;
            l.ShapeProperties = ShapeProperties.Clone();

            return l;
        }
    }
}