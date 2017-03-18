using System.Collections.Generic;
using System.Drawing;
using Ups.Toolkit.SpreadsheetLight.Core.Drawing;
using Ups.Toolkit.SpreadsheetLight.Core.misc;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Ups.Toolkit.SpreadsheetLight.Core.Charts
{
    /// <summary>
    ///     Encapsulates properties and methods for setting titles for charts.
    ///     This simulates the DocumentFormat.OpenXml.Drawing.Charts.Title class.
    /// </summary>
    public class SLTitle : SLChartAlignment
    {
        internal SLShapeProperties ShapeProperties;

        internal SLTitle(List<Color> ThemeColors, bool IsStylish = false)
        {
            // just put in the theme colors, even though it's probably not needed.
            // Memory optimisations? Take it out.
            rst = new SLRstType(SLConstants.OfficeThemeMajorLatinFont, SLConstants.OfficeThemeMinorLatinFont,
                ThemeColors, new List<Color>());
            Overlay = false;
            ShapeProperties = new SLShapeProperties(ThemeColors);

            if (IsStylish)
            {
                ShapeProperties.Fill.SetNoFill();
                ShapeProperties.Outline.SetNoLine();
            }

            RemoveTextAlignment();
        }

        internal SLRstType rst { get; set; }

        /// <summary>
        ///     Title text. This returns the plain text version if rich text is applied.
        /// </summary>
        public string Text
        {
            get { return rst.ToPlainString(); }
            set
            {
                rst = new SLRstType(SLConstants.OfficeThemeMajorLatinFont, SLConstants.OfficeThemeMinorLatinFont,
                    new List<Color>(), new List<Color>());
                rst.SetText(value);
            }
        }

        /// <summary>
        ///     Specifies if the title overlaps.
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
        ///     3D format properties.
        /// </summary>
        public SLFormat3D Format3D
        {
            get { return ShapeProperties.Format3D; }
        }

        /// <summary>
        ///     Clear all styling shape properties. Use this if you want to start styling from a clean slate.
        /// </summary>
        public void ClearShapeProperties()
        {
            ShapeProperties = new SLShapeProperties(ShapeProperties.listThemeColors);
        }

        /// <summary>
        ///     Set the title text.
        /// </summary>
        /// <param name="Text">The title text.</param>
        public void SetTitle(string Text)
        {
            this.Text = Text;
        }

        /// <summary>
        ///     Set the title with a rich text string.
        /// </summary>
        /// <param name="RichText">The rich text.</param>
        public void SetTitle(SLRstType RichText)
        {
            rst = RichText.Clone();
        }

        internal C.Title ToTitle(bool IsStylish = false)
        {
            var t = new C.Title();

            var bHasText = rst.ToPlainString().Length > 0;
            if (bHasText || (Rotation != null) || (Vertical != null) || (Anchor != null) || (AnchorCenter != null))
            {
                t.ChartText = new C.ChartText();
                t.ChartText.RichText = new C.RichText();
                t.ChartText.RichText.BodyProperties = new A.BodyProperties();

                if ((Rotation != null) || (Vertical != null) || (Anchor != null) || (AnchorCenter != null))
                {
                    if (Rotation != null)
                        t.ChartText.RichText.BodyProperties.Rotation =
                            (int) (Rotation.Value*SLConstants.DegreeToAngleRepresentation);
                    if (Vertical != null) t.ChartText.RichText.BodyProperties.Vertical = Vertical.Value;
                    if (Anchor != null) t.ChartText.RichText.BodyProperties.Anchor = Anchor.Value;
                    if (AnchorCenter != null) t.ChartText.RichText.BodyProperties.AnchorCenter = AnchorCenter.Value;
                }

                t.ChartText.RichText.ListStyle = new A.ListStyle();

                if (bHasText) t.ChartText.RichText.Append(rst.ToParagraph());
            }

            t.Layout = new C.Layout();
            t.Overlay = new C.Overlay {Val = Overlay};
            if (ShapeProperties.HasShapeProperties)
                t.ChartShapeProperties = ShapeProperties.ToChartShapeProperties(IsStylish);

            return t;
        }

        internal SLTitle Clone()
        {
            var t = new SLTitle(ShapeProperties.listThemeColors);
            t.Rotation = Rotation;
            t.Vertical = Vertical;
            t.Anchor = Anchor;
            t.AnchorCenter = AnchorCenter;
            t.rst = rst.Clone();
            t.Overlay = Overlay;
            t.ShapeProperties = ShapeProperties.Clone();

            return t;
        }
    }
}