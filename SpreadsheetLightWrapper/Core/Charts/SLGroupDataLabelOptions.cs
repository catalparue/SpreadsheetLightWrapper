using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using SpreadsheetLightWrapper.Core.misc;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLightWrapper.Core.Charts
{
    /// <summary>
    ///     Encapsulates properties and methods for setting group data label options for charts.
    /// </summary>
    public class SLGroupDataLabelOptions : EGDLblShared
    {
        // TODO Leaderlines (pie charts)

        internal SLGroupDataLabelOptions(List<Color> ThemeColors) : base(ThemeColors)
        {
            ShowLeaderLines = false;
        }

        /// <summary>
        ///     Specifies if leader lines are shown. This is for pie charts (I think...).
        /// </summary>
        public bool ShowLeaderLines { get; set; }

        internal C.DataLabels ToDataLabels(Dictionary<int, SLDataLabelOptions> Options, bool ToDelete)
        {
            var lbls = new C.DataLabels();

            if (Options.Count > 0)
            {
                var indexlist = Options.Keys.ToList();
                indexlist.Sort();
                int index;
                for (var i = 0; i < indexlist.Count; ++i)
                {
                    index = indexlist[i];
                    lbls.Append(Options[index].ToDataLabel(index));
                }
            }

            if (ToDelete)
            {
                lbls.Append(new C.Delete {Val = true});
            }
            else
            {
                if (HasNumberingFormat)
                    lbls.Append(new C.NumberingFormat {FormatCode = FormatCode, SourceLinked = SourceLinked});

                if (ShapeProperties.HasShapeProperties) lbls.Append(ShapeProperties.ToChartShapeProperties());

                if ((Rotation != null) || (Vertical != null) || (Anchor != null) || (AnchorCenter != null))
                {
                    var txtprops = new C.TextProperties();
                    txtprops.BodyProperties = new A.BodyProperties();
                    if (Rotation != null)
                        txtprops.BodyProperties.Rotation =
                            (int) (Rotation.Value*SLConstants.DegreeToAngleRepresentation);
                    if (Vertical != null) txtprops.BodyProperties.Vertical = Vertical.Value;
                    if (Anchor != null) txtprops.BodyProperties.Anchor = Anchor.Value;
                    if (AnchorCenter != null) txtprops.BodyProperties.AnchorCenter = AnchorCenter.Value;

                    txtprops.ListStyle = new A.ListStyle();

                    var para = new A.Paragraph();
                    para.ParagraphProperties = new A.ParagraphProperties();
                    para.ParagraphProperties.Append(new A.DefaultRunProperties());
                    txtprops.Append(para);

                    lbls.Append(txtprops);
                }

                if (vLabelPosition != null) lbls.Append(new C.DataLabelPosition {Val = vLabelPosition.Value});

                lbls.Append(new C.ShowLegendKey {Val = ShowLegendKey});
                lbls.Append(new C.ShowValue {Val = ShowValue});
                lbls.Append(new C.ShowCategoryName {Val = ShowCategoryName});
                lbls.Append(new C.ShowSeriesName {Val = ShowSeriesName});
                lbls.Append(new C.ShowPercent {Val = ShowPercentage});
                lbls.Append(new C.ShowBubbleSize {Val = ShowBubbleSize});

                if ((Separator != null) && (Separator.Length > 0)) lbls.Append(new C.Separator {Text = Separator});

                if (ShowLeaderLines) lbls.Append(new C.ShowLeaderLines {Val = ShowLeaderLines});
            }

            return lbls;
        }

        internal SLGroupDataLabelOptions Clone()
        {
            var gdlo = new SLGroupDataLabelOptions(ShapeProperties.listThemeColors);
            gdlo.Rotation = Rotation;
            gdlo.Vertical = Vertical;
            gdlo.Anchor = Anchor;
            gdlo.AnchorCenter = AnchorCenter;
            gdlo.HasNumberingFormat = HasNumberingFormat;
            gdlo.sFormatCode = sFormatCode;
            gdlo.bSourceLinked = bSourceLinked;
            gdlo.vLabelPosition = vLabelPosition;
            gdlo.ShapeProperties = ShapeProperties.Clone();
            gdlo.ShowLegendKey = ShowLegendKey;
            gdlo.ShowValue = ShowValue;
            gdlo.ShowCategoryName = ShowCategoryName;
            gdlo.ShowSeriesName = ShowSeriesName;
            gdlo.ShowPercentage = ShowPercentage;
            gdlo.ShowBubbleSize = ShowBubbleSize;
            gdlo.Separator = Separator;
            gdlo.ShowLeaderLines = ShowLeaderLines;

            return gdlo;
        }
    }
}