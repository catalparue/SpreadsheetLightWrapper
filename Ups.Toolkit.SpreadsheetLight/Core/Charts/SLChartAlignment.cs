using Ups.Toolkit.SpreadsheetLight.Core.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace Ups.Toolkit.SpreadsheetLight.Core.Charts
{
    /// <summary>
    ///     Encapsulates properties and methods for setting alignment in charts.
    /// </summary>
    public abstract class SLChartAlignment
    {
        /// <summary>
        ///     Initializes an instance of SLChartAlignment.
        /// </summary>
        public SLChartAlignment()
        {
            RemoveTextAlignment();
        }

        internal decimal? Rotation { get; set; }
        internal A.TextVerticalValues? Vertical { get; set; }
        internal A.TextAnchoringTypeValues? Anchor { get; set; }
        internal bool? AnchorCenter { get; set; }

        /// <summary>
        ///     Set a horizontal text direction.
        /// </summary>
        /// <param name="TextAlignment">The vertical text alignment in horizontal direction.</param>
        /// <param name="CustomAngle">Rotation angle, ranging from -90 to 90 degrees. Accurate to 1/60000 of a degree.</param>
        public void SetHorizontalTextDirection(SLTextVerticalAlignment TextAlignment, decimal CustomAngle)
        {
            if (CustomAngle < -90m) CustomAngle = -90m;
            if (CustomAngle > 90m) CustomAngle = 90m;

            // vertical axis having 0 degrees won't have the text horizontal.
            // So don't set null?
            //if (CustomAngle == 0m) this.Rotation = null;
            //else this.Rotation = CustomAngle;

            //if (CustomAngle == 0m) this.Vertical = null;
            //else this.Vertical = A.TextVerticalValues.Horizontal;

            Rotation = CustomAngle;
            Vertical = A.TextVerticalValues.Horizontal;

            switch (TextAlignment)
            {
                case SLTextVerticalAlignment.Top:
                    Anchor = A.TextAnchoringTypeValues.Top;
                    AnchorCenter = false;
                    break;
                case SLTextVerticalAlignment.Middle:
                    Anchor = A.TextAnchoringTypeValues.Center;
                    AnchorCenter = false;
                    break;
                case SLTextVerticalAlignment.Bottom:
                    Anchor = A.TextAnchoringTypeValues.Bottom;
                    AnchorCenter = false;
                    break;
                case SLTextVerticalAlignment.TopCentered:
                    Anchor = A.TextAnchoringTypeValues.Top;
                    AnchorCenter = true;
                    break;
                case SLTextVerticalAlignment.MiddleCentered:
                    Anchor = A.TextAnchoringTypeValues.Center;
                    AnchorCenter = true;
                    break;
                case SLTextVerticalAlignment.BottomCentered:
                    Anchor = A.TextAnchoringTypeValues.Bottom;
                    AnchorCenter = true;
                    break;
            }
        }

        /// <summary>
        ///     Set a stacked (vertical) text direction.
        /// </summary>
        /// <param name="TextAlignment">The horizontal text alignment in vertical direction.</param>
        /// <param name="LeftToRight">True if the text runs left-to-right. False if the text runs right-to-left.</param>
        public void SetStackedTextDirection(SLTextHorizontalAlignment TextAlignment, bool LeftToRight)
        {
            Rotation = 0m;

            Vertical = LeftToRight ? A.TextVerticalValues.WordArtVertical : A.TextVerticalValues.WordArtLeftToRight;

            switch (TextAlignment)
            {
                case SLTextHorizontalAlignment.Left:
                    if (LeftToRight)
                    {
                        Anchor = A.TextAnchoringTypeValues.Top;
                        AnchorCenter = false;
                    }
                    else
                    {
                        Anchor = A.TextAnchoringTypeValues.Bottom;
                        AnchorCenter = false;
                    }
                    break;
                case SLTextHorizontalAlignment.Center:
                    Anchor = A.TextAnchoringTypeValues.Center;
                    AnchorCenter = false;
                    break;
                case SLTextHorizontalAlignment.Right:
                    if (LeftToRight)
                    {
                        Anchor = A.TextAnchoringTypeValues.Bottom;
                        AnchorCenter = false;
                    }
                    else
                    {
                        Anchor = A.TextAnchoringTypeValues.Top;
                        AnchorCenter = false;
                    }
                    break;
                case SLTextHorizontalAlignment.LeftMiddle:
                    if (LeftToRight)
                    {
                        Anchor = A.TextAnchoringTypeValues.Top;
                        AnchorCenter = false;
                    }
                    else
                    {
                        Anchor = A.TextAnchoringTypeValues.Bottom;
                        AnchorCenter = false;
                    }
                    break;
                case SLTextHorizontalAlignment.CenterMiddle:
                    Anchor = A.TextAnchoringTypeValues.Center;
                    AnchorCenter = true;
                    break;
                case SLTextHorizontalAlignment.RightMiddle:
                    if (LeftToRight)
                    {
                        Anchor = A.TextAnchoringTypeValues.Bottom;
                        AnchorCenter = true;
                    }
                    else
                    {
                        Anchor = A.TextAnchoringTypeValues.Top;
                        AnchorCenter = true;
                    }
                    break;
            }
        }

        /// <summary>
        ///     Set the text rotated 90 degrees.
        /// </summary>
        public void SetTextRotated90Degrees()
        {
            Rotation = 90m;
            Vertical = A.TextVerticalValues.Horizontal;
            Anchor = A.TextAnchoringTypeValues.Top;
            AnchorCenter = false;
        }

        /// <summary>
        ///     Set the text rotated 270 degrees.
        /// </summary>
        public void SetTextRotated270Degrees()
        {
            Rotation = -90m;
            Vertical = A.TextVerticalValues.Horizontal;
            Anchor = A.TextAnchoringTypeValues.Top;
            AnchorCenter = false;
        }

        /// <summary>
        ///     Remove all text alignment.
        /// </summary>
        public void RemoveTextAlignment()
        {
            Rotation = null;
            Vertical = null;
            Anchor = null;
            AnchorCenter = null;
        }
    }
}