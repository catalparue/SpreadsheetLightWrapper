using System.Collections.Generic;
using System.Drawing;
using Ups.Toolkit.SpreadsheetLight.Core.misc;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Ups.Toolkit.SpreadsheetLight.Core.Charts
{
    /// <summary>
    ///     Encapsulates properties and methods for setting data label options for charts.
    /// </summary>
    public class SLDataLabelOptions : EGDLblShared
    {
        internal SLDataLabelOptions(List<Color> ThemeColors) : base(ThemeColors)
        {
            RichText = null;
        }

        // TODO Layout?

        internal SLRstType RichText { get; set; }

        /// <summary>
        ///     Set custom label text.
        /// </summary>
        /// <param name="Text">The custom text.</param>
        public void SetLabelText(string Text)
        {
            var rst = new SLRstType();
            rst.AppendText(Text);
            RichText = rst.Clone();
        }

        /// <summary>
        ///     Set custom label text.
        /// </summary>
        /// <param name="RichText">The custom text in rich text format.</param>
        public void SetLabelText(SLRstType RichText)
        {
            this.RichText = RichText.Clone();
        }

        /// <summary>
        ///     Reset the label text. This removes any custom label text.
        /// </summary>
        public void ResetLabelText()
        {
            RichText = null;
        }

        internal C.DataLabel ToDataLabel(int index)
        {
            var lbl = new C.DataLabel();

            lbl.Index = new C.Index {Val = (uint) index};

            lbl.Append(new C.Layout());

            if ((RichText != null) || (Rotation != null) || (Vertical != null) || (Anchor != null) ||
                (AnchorCenter != null))
            {
                var ctxt = new C.ChartText();
                ctxt.RichText = new C.RichText();
                ctxt.RichText.BodyProperties = new A.BodyProperties();

                if ((Rotation != null) || (Vertical != null) || (Anchor != null) || (AnchorCenter != null))
                {
                    if (Rotation != null)
                        ctxt.RichText.BodyProperties.Rotation =
                            (int) (Rotation.Value*SLConstants.DegreeToAngleRepresentation);
                    if (Vertical != null) ctxt.RichText.BodyProperties.Vertical = Vertical.Value;
                    if (Anchor != null) ctxt.RichText.BodyProperties.Anchor = Anchor.Value;
                    if (AnchorCenter != null) ctxt.RichText.BodyProperties.AnchorCenter = AnchorCenter.Value;
                }

                ctxt.RichText.ListStyle = new A.ListStyle();

                if (RichText != null) ctxt.RichText.Append(RichText.ToParagraph());

                lbl.Append(ctxt);
            }

            if (HasNumberingFormat)
                lbl.Append(new C.NumberingFormat {FormatCode = FormatCode, SourceLinked = SourceLinked});

            if (ShapeProperties.HasShapeProperties) lbl.Append(ShapeProperties.ToChartShapeProperties());

            if (vLabelPosition != null) lbl.Append(new C.DataLabelPosition {Val = vLabelPosition.Value});

            lbl.Append(new C.ShowLegendKey {Val = ShowLegendKey});
            lbl.Append(new C.ShowValue {Val = ShowValue});
            lbl.Append(new C.ShowCategoryName {Val = ShowCategoryName});
            lbl.Append(new C.ShowSeriesName {Val = ShowSeriesName});
            lbl.Append(new C.ShowPercent {Val = ShowPercentage});
            lbl.Append(new C.ShowBubbleSize {Val = ShowBubbleSize});

            if ((Separator != null) && (Separator.Length > 0)) lbl.Append(new C.Separator {Text = Separator});

            return lbl;
        }

        internal SLDataLabelOptions Clone()
        {
            var dlo = new SLDataLabelOptions(ShapeProperties.listThemeColors);
            dlo.Rotation = Rotation;
            dlo.Vertical = Vertical;
            dlo.Anchor = Anchor;
            dlo.AnchorCenter = AnchorCenter;
            dlo.HasNumberingFormat = HasNumberingFormat;
            dlo.sFormatCode = sFormatCode;
            dlo.bSourceLinked = bSourceLinked;
            dlo.vLabelPosition = vLabelPosition;
            dlo.ShapeProperties = ShapeProperties.Clone();
            dlo.ShowLegendKey = ShowLegendKey;
            dlo.ShowValue = ShowValue;
            dlo.ShowCategoryName = ShowCategoryName;
            dlo.ShowSeriesName = ShowSeriesName;
            dlo.ShowPercentage = ShowPercentage;
            dlo.ShowBubbleSize = ShowBubbleSize;
            dlo.Separator = Separator;
            if (RichText != null) dlo.RichText = RichText.Clone();

            return dlo;
        }
    }
}