using System.Collections.Generic;
using System.Drawing;
using DocumentFormat.OpenXml.Vml;
using SpreadsheetLightWrapper.Core.misc;
using SpreadsheetLightWrapper.Core.style;
using SLFill = SpreadsheetLightWrapper.Core.Drawing.SLFill;

namespace SpreadsheetLightWrapper.Core.worksheet
{
    /// <summary>
    ///     Specifies how the text is aligned horizontally.
    /// </summary>
    public enum SLHorizontalTextAlignmentValues
    {
        /// <summary>
        ///     Left
        /// </summary>
        Left = 0,

        /// <summary>
        ///     Justify
        /// </summary>
        Justify,

        /// <summary>
        ///     Center
        /// </summary>
        Center,

        /// <summary>
        ///     Right
        /// </summary>
        Right,

        /// <summary>
        ///     Distributed
        /// </summary>
        Distributed
    }

    /// <summary>
    ///     Specifies how the text is aligned vertically.
    /// </summary>
    public enum SLVerticalTextAlignmentValues
    {
        /// <summary>
        ///     Top
        /// </summary>
        Top = 0,

        /// <summary>
        ///     Justify
        /// </summary>
        Justify,

        /// <summary>
        ///     Center
        /// </summary>
        Center,

        /// <summary>
        ///     Bottom
        /// </summary>
        Bottom,

        /// <summary>
        ///     Distributed
        /// </summary>
        Distributed
    }

    /// <summary>
    ///     Specifies how the comment is oriented.
    /// </summary>
    public enum SLCommentOrientationValues
    {
        /// <summary>
        ///     Horizontal
        /// </summary>
        Horizontal = 0,

        /// <summary>
        ///     The text characters are arranged in a top-down direction
        /// </summary>
        TopDown,

        /// <summary>
        ///     Rotated 270 degrees
        /// </summary>
        Rotated270Degrees,

        /// <summary>
        ///     Rotated 90 degrees
        /// </summary>
        Rotated90Degrees
    }

    /// <summary>
    ///     Specifies how line dashes are styled
    /// </summary>
    public enum SLDashStyleValues
    {
        /// <summary>
        ///     Solid
        /// </summary>
        Solid = 0,

        /// <summary>
        ///     Short dash
        /// </summary>
        ShortDash,

        /// <summary>
        ///     Short dot
        /// </summary>
        ShortDot,

        /// <summary>
        ///     Short dash dot
        /// </summary>
        ShortDashDot,

        /// <summary>
        ///     Short dash dot dot
        /// </summary>
        ShortDashDotDot,

        /// <summary>
        ///     Dot
        /// </summary>
        Dot,

        /// <summary>
        ///     Dash
        /// </summary>
        Dash,

        /// <summary>
        ///     Long dash
        /// </summary>
        LongDash,

        /// <summary>
        ///     Dash dot
        /// </summary>
        DashDot,

        /// <summary>
        ///     Long dash dot
        /// </summary>
        LongDashDot,

        /// <summary>
        ///     Long dash dot dot
        /// </summary>
        LongDashDotDot
    }

    /// <summary>
    ///     Encapsulates properties and methods for cell comments.
    /// </summary>
    public class SLComment
    {
        internal byte bFromTransparency;
        internal byte bToTransparency;

        internal double fHeight;

        internal double? fLineWeight;

        internal double fWidth;

        internal bool HasSetPosition;
        internal List<Color> listThemeColors;

        internal SLRstType rst;

        // TODO: move with cells and size with cells

        internal string sAuthor;

        internal bool UsePositionMargin;
        internal StrokeEndCapValues? vEndCap;

        internal SLDashStyleValues? vLineDashStyle;

        internal SLComment(List<Color> ThemeColors)
        {
            int i;
            listThemeColors = new List<Color>();
            for (i = 0; i < ThemeColors.Count; ++i)
                listThemeColors.Add(ThemeColors[i]);

            SetAllNull();
        }

        /// <summary>
        ///     The author of the comment.
        /// </summary>
        public string Author
        {
            get { return sAuthor; }
            set { sAuthor = value.Trim(); }
        }

        internal double Top { get; set; }
        internal double Left { get; set; }
        internal double TopMargin { get; set; }
        internal double LeftMargin { get; set; }

        /// <summary>
        ///     Set true to automatically size the comment box according to the comment's contents.
        /// </summary>
        public bool AutoSize { get; set; }

        /// <summary>
        ///     Width of comment box in units of points. For practical purposes, the width is a minimum of 1 pt.
        /// </summary>
        public double Width
        {
            get { return fWidth; }
            set
            {
                fWidth = value;
                if (fWidth < 1.0) fWidth = 1.0;
                AutoSize = false;
            }
        }

        /// <summary>
        ///     Height of comment box in units of points. For practical purposes, the height is a minimum of 1 pt.
        /// </summary>
        public double Height
        {
            get { return fHeight; }
            set
            {
                fHeight = value;
                if (fHeight < 1.0) fHeight = 1.0;
                AutoSize = false;
            }
        }

        /// <summary>
        ///     Fill properties. Note that this is repurposed, and some of the methods and properties can't be
        ///     directly translated to a VML-equivalent (which is how comment styles are stored).
        /// </summary>
        public SLFill Fill { get; set; }

        /// <summary>
        ///     The transparency value of the first gradient point measured in percentage, ranging from 0% to 100% (both
        ///     inclusive).
        /// </summary>
        public byte GradientFromTransparency
        {
            get { return bFromTransparency; }
            set
            {
                bFromTransparency = value;
                if (bFromTransparency > 100) bFromTransparency = 100;
            }
        }

        /// <summary>
        ///     The transparency value of the last gradient point measured in percentage, ranging from 0% to 100% (both inclusive).
        /// </summary>
        public byte GradientToTransparency
        {
            get { return bToTransparency; }
            set
            {
                bToTransparency = value;
                if (bToTransparency > 100) bToTransparency = 100;
            }
        }

        /// <summary>
        ///     Set null for automatic color.
        /// </summary>
        public Color? LineColor { get; set; }

        /// <summary>
        ///     Line weight in points.
        /// </summary>
        public double LineWeight
        {
            // 0.75pt seems to be Excel's default, although the Open XML specs state 1pt as the default
            get { return fLineWeight ?? 0.75; }
            set
            {
                fLineWeight = value;
                if (fLineWeight < 0) fLineWeight = 0;
            }
        }

        /// <summary>
        ///     Line style.
        /// </summary>
        public StrokeLineStyleValues LineStyle { get; set; }

        /// <summary>
        ///     Horizontal text alignment.
        /// </summary>
        public SLHorizontalTextAlignmentValues HorizontalTextAlignment { get; set; }

        /// <summary>
        ///     Vertical text alignment.
        /// </summary>
        public SLVerticalTextAlignmentValues VerticalTextAlignment { get; set; }

        /// <summary>
        ///     Comment text orientation.
        /// </summary>
        public SLCommentOrientationValues Orientation { get; set; }

        /// <summary>
        ///     Comment text direction.
        /// </summary>
        public SLAlignmentReadingOrderValues TextDirection { get; set; }

        /// <summary>
        ///     Specifies whether the comment box has a shadow.
        /// </summary>
        public bool HasShadow { get; set; }

        /// <summary>
        ///     Specifies the color of the comment box's shadow.
        /// </summary>
        public Color ShadowColor { get; set; }

        /// <summary>
        ///     Specifies whether the comment is visible.
        /// </summary>
        public bool Visible { get; set; }

        private void SetAllNull()
        {
            sAuthor = string.Empty;
            rst = new SLRstType();
            HasSetPosition = false;
            Top = 0;
            Left = 0;
            UsePositionMargin = false;
            TopMargin = 0;
            LeftMargin = 0;
            AutoSize = false;
            fWidth = SLConstants.DefaultCommentBoxWidth;
            fHeight = SLConstants.DefaultCommentBoxHeight;

            Fill = new SLFill(listThemeColors);
            Fill.SetSolidFill(Color.FromArgb(255, 255, 225), 0);
            bFromTransparency = 0;
            bToTransparency = 0;

            LineColor = null;
            fLineWeight = null;
            LineStyle = StrokeLineStyleValues.Single;
            vLineDashStyle = null;
            vEndCap = null;
            HorizontalTextAlignment = SLHorizontalTextAlignmentValues.Left;
            VerticalTextAlignment = SLVerticalTextAlignmentValues.Top;
            Orientation = SLCommentOrientationValues.Horizontal;
            TextDirection = SLAlignmentReadingOrderValues.ContextDependent;

            HasShadow = true;
            ShadowColor = Color.Black;

            Visible = false;
        }

        /// <summary>
        ///     Set the comment text.
        /// </summary>
        /// <param name="Text">The comment text.</param>
        public void SetText(string Text)
        {
            rst = new SLRstType();
            rst.SetText(Text);
        }

        /// <summary>
        ///     Set the comment text given rich text content.
        /// </summary>
        /// <param name="RichText">The rich text content</param>
        public void SetText(SLRstType RichText)
        {
            rst = new SLRstType();
            rst = RichText.Clone();
        }

        /// <summary>
        ///     Set the position of the comment box. NOTE: This isn't an exact science. The positioning depends on the DPI of the
        ///     computer's screen.
        /// </summary>
        /// <param name="Top">
        ///     Top position of the comment box based on row index. For example, 0.5 means at the half-way point of
        ///     the 1st row, 2.5 means at the half-way point of the 3rd row.
        /// </param>
        /// <param name="Left">
        ///     Left position of the comment box based on column index. For example, 0.5 means at the half-way point
        ///     of the 1st column, 2.5 means at the half-way point of the 3rd column.
        /// </param>
        public void SetPosition(double Top, double Left)
        {
            HasSetPosition = true;
            this.Top = Top;
            this.Left = Left;
        }

        /// <summary>
        ///     Set the position of the comment box given the top and left margins measured in points. It is suggested to use
        ///     SetPosition() instead. This method is provided as a means of convenience. NOTE: This isn't an exact science. The
        ///     positioning depends on the DPI of the computer's screen.
        /// </summary>
        /// <param name="TopMargin">Top margin in points. This is measured from the top-left corner of the cell A1.</param>
        /// <param name="LeftMargin">Left margin in points. This is measured from the top-left corner of the cell A1.</param>
        public void SetPositionMargin(double TopMargin, double LeftMargin)
        {
            HasSetPosition = true;
            UsePositionMargin = true;
            this.TopMargin = TopMargin;
            this.LeftMargin = LeftMargin;
        }

        /// <summary>
        ///     Set the dash style of the comment box.
        /// </summary>
        /// <param name="DashStyle">The dash style.</param>
        public void SetDashStyle(SLDashStyleValues DashStyle)
        {
            vLineDashStyle = DashStyle;
            vEndCap = null;
        }

        /// <summary>
        ///     Set the dash style of the comment box.
        /// </summary>
        /// <param name="DashStyle">The dash style.</param>
        /// <param name="EndCap">The end cap of the lines.</param>
        public void SetDashStyle(SLDashStyleValues DashStyle, StrokeEndCapValues EndCap)
        {
            vLineDashStyle = DashStyle;
            vEndCap = EndCap;
        }

        internal SLComment Clone()
        {
            var comm = new SLComment(listThemeColors);
            comm.sAuthor = sAuthor;
            comm.rst = rst.Clone();
            comm.HasSetPosition = HasSetPosition;
            comm.Top = Top;
            comm.Left = Left;
            comm.UsePositionMargin = UsePositionMargin;
            comm.TopMargin = TopMargin;
            comm.LeftMargin = LeftMargin;
            comm.AutoSize = AutoSize;
            comm.fWidth = fWidth;
            comm.fHeight = fHeight;
            comm.Fill = Fill.Clone();
            comm.bFromTransparency = bFromTransparency;
            comm.bToTransparency = bToTransparency;
            comm.LineColor = LineColor;
            comm.fLineWeight = fLineWeight;
            comm.LineStyle = LineStyle;
            comm.vLineDashStyle = vLineDashStyle;
            comm.vEndCap = vEndCap;
            comm.HorizontalTextAlignment = HorizontalTextAlignment;
            comm.VerticalTextAlignment = VerticalTextAlignment;
            comm.Orientation = Orientation;
            comm.TextDirection = TextDirection;
            comm.HasShadow = HasShadow;
            comm.ShadowColor = ShadowColor;
            comm.Visible = Visible;

            return comm;
        }
    }
}