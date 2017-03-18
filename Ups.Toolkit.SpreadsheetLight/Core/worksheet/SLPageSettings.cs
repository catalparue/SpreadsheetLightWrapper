using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;
using Ups.Toolkit.SpreadsheetLight.Core.misc;
using Ups.Toolkit.SpreadsheetLight.Core.style;
using Color = System.Drawing.Color;

namespace Ups.Toolkit.SpreadsheetLight.Core.worksheet
{
    internal enum SLHeaderFooterSection
    {
        None = 0,
        OddHeader,
        OddFooter,
        EvenHeader,
        EvenFooter,
        FirstHeader,
        FirstFooter
    }

    /// <summary>
    ///     Encapsulates page and print settings for a sheet (worksheets, chartsheets and dialogsheets).
    ///     This simulates DocumentFormat.OpenXml.Spreadsheet.SheetProperties, DocumentFormat.OpenXml.Spreadsheet.PrintOptions,
    ///     DocumentFormat.OpenXml.Spreadsheet.PageMargins,
    ///     DocumentFormat.OpenXml.Spreadsheet.PageSetup, DocumentFormat.OpenXml.Spreadsheet.HeaderFooter and
    ///     DocumentFormat.OpenXml.Spreadsheet.SheetView classes.
    ///     For chartsheets, the DocumentFormat.OpenXml.Spreadsheet.ChartSheetProperties (instead of SheetProperties) and
    ///     DocumentFormat.OpenXml.Spreadsheet.ChartSheetPageSetup (instead of PageSetup) classes are involved.
    /// </summary>
    public class SLPageSettings
    {
        // Apparently when show formulas is true, column widths are doubled. Erhmahgerd...

        internal bool? bShowFormulas;

        internal bool? bShowGridLines;

        internal bool? bShowRowColumnHeaders;

        internal bool? bShowRuler;

        internal double fBottomMargin;

        internal double fFooterMargin;

        internal double fHeaderMargin;

        internal double fLeftMargin;

        internal double fRightMargin;

        internal double fTopMargin;

        internal bool HasPageMargins;

        internal uint iCopies;

        internal uint iFirstPageNumber;

        internal uint iFitToHeight;

        internal uint iFitToWidth;

        internal uint iScale;

        internal uint? iZoomScale;

        internal uint? iZoomScaleNormal;

        internal uint? iZoomScalePageLayoutView;
        internal List<Color> listIndexedColors;

        internal List<Color> listThemeColors;

        internal SheetViewValues? vView;

        /// <summary>
        ///     Initializes an instance of SLPageSettings. It is recommended to use GetPageSettings() of the SLDocument class.
        /// </summary>
        public SLPageSettings()
        {
            Initialize(new List<Color>(), new List<Color>());
        }

        internal SLPageSettings(List<Color> ThemeColors, List<Color> IndexedColors)
        {
            Initialize(ThemeColors, IndexedColors);
        }

        //SheetProperties: TabColor
        //SheetProperties: PageSetupProperties
        //SheetViews: Zoom
        //PrintOptions (parents: customSheetView, dialogsheet, worksheet)
        //PageMargins (parents: chartsheet, customSheetView, dialogsheet, worksheet)
        //PageSetup (parents: customSheetView, dialogsheet, worksheet)
        //HeaderFooter (parents: chartsheet, customSheetView, dialogsheet, worksheet)

        internal bool HasSheetProperties
        {
            get { return SheetProperties.HasSheetProperties; }
        }

        internal bool HasChartSheetProperties
        {
            get { return SheetProperties.HasChartSheetProperties; }
        }

        internal SLSheetProperties SheetProperties { get; set; }

        /// <summary>
        ///     Specifies if there's a tab color. This is read-only.
        /// </summary>
        public bool HasTabColor
        {
            get { return SheetProperties.HasTabColor; }
        }

        /// <summary>
        ///     The tab color.
        /// </summary>
        public Color TabColor
        {
            get { return SheetProperties.clrTabColor.Color; }
            set { SheetProperties.TabColor = value; }
        }

        internal bool HasSheetView
        {
            get
            {
                return (bShowFormulas != null) || (bShowGridLines != null)
                       || (bShowRowColumnHeaders != null) || (bShowRuler != null)
                       || (vView != null) || (iZoomScale != null)
                       || (iZoomScaleNormal != null) || (iZoomScalePageLayoutView != null);
            }
        }

        /// <summary>
        ///     Show or hide the cell formulas. NOTE: This has nothing to do with the formula bar, but whether the sheet shows cell
        ///     formulas instead of calculated results.
        /// </summary>
        public bool ShowFormulas
        {
            get { return bShowFormulas ?? false; }
            set { bShowFormulas = value; }
        }

        /// <summary>
        ///     Show or hide the grid lines between rows and columns.
        /// </summary>
        public bool ShowGridLines
        {
            get { return bShowGridLines ?? true; }
            set { bShowGridLines = value; }
        }

        /// <summary>
        ///     Show or hide the row and column headers.
        /// </summary>
        public bool ShowRowColumnHeaders
        {
            get { return bShowRowColumnHeaders ?? true; }
            set { bShowRowColumnHeaders = value; }
        }

        /// <summary>
        ///     Show or hide the ruler on the worksheet. The ruler is only seen when the worksheet view is in "page layout" mode.
        /// </summary>
        public bool ShowRuler
        {
            get { return bShowRuler ?? true; }
            set { bShowRuler = value; }
        }

        /// <summary>
        ///     Worksheet view type.
        /// </summary>
        public SheetViewValues View
        {
            get { return vView ?? SheetViewValues.Normal; }
            set { vView = value; }
        }

        /// <summary>
        ///     Zoom magnification for current view, ranging from 10% to 400%. If you want to set a zoom value for the page break
        ///     view, make sure to set the View property to PageBreakPreview.
        /// </summary>
        public uint ZoomScale
        {
            get { return iZoomScale ?? 100; }
            set
            {
                iZoomScale = value;
                if (iZoomScale < 10) iZoomScale = 10;
                if (iZoomScale > 400) iZoomScale = 400;
            }
        }

        /// <summary>
        ///     Zoom magnification for the normal view, ranging from 10% to 400%. A return value of 0% means the automatic setting
        ///     is used.
        ///     If the view is set to normal, this value is ignored if ZoomScale is also set.
        /// </summary>
        public uint ZoomScaleNormal
        {
            get { return iZoomScaleNormal ?? 0; }
            set
            {
                iZoomScaleNormal = value;
                if (iZoomScaleNormal < 10) iZoomScaleNormal = 10;
                if (iZoomScaleNormal > 400) iZoomScaleNormal = 400;
            }
        }

        /// <summary>
        ///     Zoom magnification for the page layout view, ranging from 10% to 400%. A return value of 0% means the automatic
        ///     setting is used.
        ///     If the view is set to page layout, this value is ignored if ZoomScale is also set.
        /// </summary>
        public uint ZoomScalePageLayoutView
        {
            get { return iZoomScalePageLayoutView ?? 0; }
            set
            {
                iZoomScalePageLayoutView = value;
                if (iZoomScalePageLayoutView < 10) iZoomScalePageLayoutView = 10;
                if (iZoomScalePageLayoutView > 400) iZoomScalePageLayoutView = 400;
            }
        }

        internal bool HasPrintOptions
        {
            get { return PrintHorizontalCentered || PrintVerticalCentered || PrintHeadings || PrintGridLines; }
        }

        /// <summary>
        ///     Center horizontally on page when printing. This doesn't apply to chart sheets.
        /// </summary>
        public bool PrintHorizontalCentered { get; set; }

        /// <summary>
        ///     Center vertically on page when printing. This doesn't apply to chart sheets.
        /// </summary>
        public bool PrintVerticalCentered { get; set; }

        /// <summary>
        ///     Print row and column headings. This doesn't apply to chart sheets.
        /// </summary>
        public bool PrintHeadings { get; set; }

        /// <summary>
        ///     Print grid lines. This doesn't apply to chart sheets.
        /// </summary>
        public bool PrintGridLines { get; set; }

        internal bool PrintGridLinesSet { get; set; }

        /// <summary>
        ///     The left margin in inches.
        /// </summary>
        public double LeftMargin
        {
            get { return fLeftMargin; }
            set
            {
                fLeftMargin = value;
                if (fLeftMargin < 0) fLeftMargin = 0;
                HasPageMargins = true;
            }
        }

        /// <summary>
        ///     The right margin in inches.
        /// </summary>
        public double RightMargin
        {
            get { return fRightMargin; }
            set
            {
                fRightMargin = value;
                if (fRightMargin < 0) fRightMargin = 0;
                HasPageMargins = true;
            }
        }

        /// <summary>
        ///     The top margin in inches.
        /// </summary>
        public double TopMargin
        {
            get { return fTopMargin; }
            set
            {
                fTopMargin = value;
                if (fTopMargin < 0) fTopMargin = 0;
                HasPageMargins = true;
            }
        }

        /// <summary>
        ///     The bottom margin in inches.
        /// </summary>
        public double BottomMargin
        {
            get { return fBottomMargin; }
            set
            {
                fBottomMargin = value;
                if (fBottomMargin < 0) fBottomMargin = 0;
                HasPageMargins = true;
            }
        }

        /// <summary>
        ///     The header margin in inches.
        /// </summary>
        public double HeaderMargin
        {
            get { return fHeaderMargin; }
            set
            {
                fHeaderMargin = value;
                if (fHeaderMargin < 0) fHeaderMargin = 0;
                HasPageMargins = true;
            }
        }

        /// <summary>
        ///     The footer margin in inches.
        /// </summary>
        public double FooterMargin
        {
            get { return fFooterMargin; }
            set
            {
                if (fFooterMargin < 0) fFooterMargin = 0;
                HasPageMargins = true;
            }
        }

        internal bool HasPageSetup
        {
            get
            {
                return (PaperSize != SLPaperSizeValues.LetterPaper) || (FirstPageNumber != 1)
                       || (Scale != 100) || (FitToWidth != 1) || (FitToHeight != 1)
                       || (PageOrder != PageOrderValues.DownThenOver) || (Orientation != OrientationValues.Default)
                       || !UsePrinterDefaults
                       || BlackAndWhite || Draft || (CellComments != CellCommentsValues.None)
                       || (Errors != PrintErrorValues.Displayed) || (HorizontalDpi != 600)
                       || (VerticalDpi != 600) || (Copies != 1);
            }
        }

        internal bool HasChartSheetPageSetup
        {
            get
            {
                return (PaperSize != SLPaperSizeValues.LetterPaper) || (FirstPageNumber != 1)
                       || (Orientation != OrientationValues.Default)
                       || !UsePrinterDefaults
                       || BlackAndWhite || Draft
                       || (HorizontalDpi != 600)
                       || (VerticalDpi != 600) || (Copies != 1);
            }
        }

        /// <summary>
        ///     The paper size. The default is Letter.
        /// </summary>
        public SLPaperSizeValues PaperSize { get; set; }

        /// <summary>
        ///     The page number set for the first printed page.
        /// </summary>
        public uint FirstPageNumber
        {
            get { return iFirstPageNumber; }
            set
            {
                iFirstPageNumber = value;
                if (iFirstPageNumber < 1) iFirstPageNumber = 1;
            }
        }

        /// <summary>
        ///     The printing scale. This is read-only. This doesn't apply to chart sheets.
        /// </summary>
        public uint Scale
        {
            get { return iScale; }
        }

        /// <summary>
        ///     The number of horizontal pages to fit into a printed page. This is read-only. This doesn't apply to chart sheets.
        /// </summary>
        public uint FitToWidth
        {
            get { return iFitToWidth; }
        }

        /// <summary>
        ///     The number of vertical pages to fit into a printed page. This is read-only. This doesn't apply to chart sheets.
        /// </summary>
        public uint FitToHeight
        {
            get { return iFitToHeight; }
        }

        /// <summary>
        ///     Page order when printed. This doesn't apply to chart sheets.
        /// </summary>
        public PageOrderValues PageOrder { get; set; }

        /// <summary>
        ///     Page orientation.
        /// </summary>
        public OrientationValues Orientation { get; set; }

        internal bool UsePrinterDefaults { get; set; }

        /// <summary>
        ///     Specifies if the page is printed in black and white.
        /// </summary>
        public bool BlackAndWhite { get; set; }

        /// <summary>
        ///     Specifies if the page is printed in draft mode (without graphics).
        /// </summary>
        public bool Draft { get; set; }

        /// <summary>
        ///     Specifies how to print cell comments. This doesn't apply to chart sheets.
        /// </summary>
        public CellCommentsValues CellComments { get; set; }

        /// <summary>
        ///     Specifies how to print for cells with errors. This doesn't apply to chart sheets.
        /// </summary>
        public PrintErrorValues Errors { get; set; }

        /// <summary>
        ///     Horizontal print resolution.
        /// </summary>
        public uint HorizontalDpi { get; set; }

        /// <summary>
        ///     Vertical print resolution.
        /// </summary>
        public uint VerticalDpi { get; set; }

        /// <summary>
        ///     The number of copies to print. The minimum number is 1 copy. There are no maximum number of copies, however Excel
        ///     uses 9999 copies as a maximum.
        /// </summary>
        public uint Copies
        {
            get { return iCopies; }
            set
            {
                iCopies = value;
                if (iCopies < 1) iCopies = 1;
            }
        }

        internal bool HasHeaderFooter
        {
            get
            {
                return (OddHeaderText.Length > 0) || (OddFooterText.Length > 0) || (EvenHeaderText.Length > 0)
                       || (EvenFooterText.Length > 0) || (FirstHeaderText.Length > 0) || (FirstFooterText.Length > 0)
                       || DifferentOddEvenPages || DifferentFirstPage || !ScaleWithDocument || !AlignWithMargins;
            }
        }

        /// <summary>
        ///     The text in the odd-numbered page header. Note that this is the default used.
        /// </summary>
        public string OddHeaderText { get; set; }

        /// <summary>
        ///     The text in the odd-numbered page footer. Note that this is the default used.
        /// </summary>
        public string OddFooterText { get; set; }

        /// <summary>
        ///     The text in the even-numbered page header. Note that this only activates when <see cref="DifferentOddEvenPages" />
        ///     is true.
        /// </summary>
        public string EvenHeaderText { get; set; }

        /// <summary>
        ///     The text in the even-numbered page footer. Note that this only activates when <see cref="DifferentOddEvenPages" />
        ///     is true.
        /// </summary>
        public string EvenFooterText { get; set; }

        /// <summary>
        ///     The text in the first page's header. Note that this only activates when <see cref="DifferentFirstPage" /> is true.
        /// </summary>
        public string FirstHeaderText { get; set; }

        /// <summary>
        ///     The text in the first page's footer. Note that this only activates when <see cref="DifferentFirstPage" /> is true.
        /// </summary>
        public string FirstFooterText { get; set; }

        /// <summary>
        ///     Specifies if different headers and footers are set for odd- and even-numbered pages.
        ///     If false, then the text in odd-numbered page header and footer is used, even if there's text
        ///     set in even-numbered page header and footer.
        /// </summary>
        public bool DifferentOddEvenPages { get; set; }

        /// <summary>
        ///     Specifies if a different header and footer is set for the first page.
        ///     If false, any text set in the first page header and footer is ignored.
        /// </summary>
        public bool DifferentFirstPage { get; set; }

        /// <summary>
        ///     Scale with the document.
        /// </summary>
        public bool ScaleWithDocument { get; set; }

        /// <summary>
        ///     Align header and footer margins with page margins.
        /// </summary>
        public bool AlignWithMargins { get; set; }

        private bool StrikeSwitch { get; set; }
        private bool SuperscriptSwitch { get; set; }
        private bool SubscriptSwitch { get; set; }
        private bool UnderlineSwitch { get; set; }
        private bool DoubleUnderlineSwitch { get; set; }
        private bool FontSizeSwitch { get; set; }
        private bool FontColorSwitch { get; set; }
        private bool FontStyleSwitch { get; set; }
        private SLHeaderFooterSection HFSection { get; set; }

        private void Initialize(List<Color> ThemeColors, List<Color> IndexedColors)
        {
            int i;
            listThemeColors = new List<Color>();
            for (i = 0; i < ThemeColors.Count; ++i)
                listThemeColors.Add(ThemeColors[i]);

            listIndexedColors = new List<Color>();
            for (i = 0; i < IndexedColors.Count; ++i)
                listIndexedColors.Add(IndexedColors[i]);

            SetAllNull();
            ResetSwitches();
        }

        private void SetAllNull()
        {
            SheetProperties = new SLSheetProperties(listThemeColors, listIndexedColors);

            bShowFormulas = null;
            bShowGridLines = null;
            bShowRowColumnHeaders = null;
            bShowRuler = null;
            vView = null;
            iZoomScale = null;
            iZoomScaleNormal = null;
            iZoomScalePageLayoutView = null;

            PrintHorizontalCentered = false;
            PrintVerticalCentered = false;
            PrintHeadings = false;
            PrintGridLines = false;
            PrintGridLinesSet = true;

            SetNormalMargins();
            HasPageMargins = false;

            PaperSize = SLPaperSizeValues.LetterPaper;
            FirstPageNumber = 1;
            iScale = 100;
            iFitToWidth = 1;
            iFitToHeight = 1;
            PageOrder = PageOrderValues.DownThenOver;
            Orientation = OrientationValues.Default;
            UsePrinterDefaults = true;
            BlackAndWhite = false;
            Draft = false;
            CellComments = CellCommentsValues.None;
            Errors = PrintErrorValues.Displayed;
            HorizontalDpi = 600;
            VerticalDpi = 600;
            Copies = 1;

            OddHeaderText = string.Empty;
            OddFooterText = string.Empty;
            EvenHeaderText = string.Empty;
            EvenFooterText = string.Empty;
            FirstHeaderText = string.Empty;
            FirstFooterText = string.Empty;
            DifferentOddEvenPages = false;
            DifferentFirstPage = false;
            ScaleWithDocument = true;
            AlignWithMargins = true;
        }

        private void ResetSwitches()
        {
            StrikeSwitch = false;
            SuperscriptSwitch = false;
            SubscriptSwitch = false;
            UnderlineSwitch = false;
            DoubleUnderlineSwitch = false;
            FontSizeSwitch = false;
            FontColorSwitch = false;
            FontStyleSwitch = false;
            HFSection = SLHeaderFooterSection.None;
        }

        /// <summary>
        ///     Sets the tab color of the sheet.
        /// </summary>
        /// <param name="TabColor">The theme color to be used.</param>
        public void SetTabColor(SLThemeColorIndexValues TabColor)
        {
            SheetProperties.clrTabColor.SetThemeColor(TabColor);
            SheetProperties.HasTabColor = SheetProperties.clrTabColor.Color.IsEmpty ? false : true;
        }

        /// <summary>
        ///     Sets the tab color of the sheet.
        /// </summary>
        /// <param name="TabColor">The theme color to be used.</param>
        /// <param name="Tint">
        ///     The tint applied to the theme color, ranging from -1.0 to 1.0. Negative tints darken the theme color
        ///     and positive tints lighten the theme color.
        /// </param>
        public void SetTabColor(SLThemeColorIndexValues TabColor, double Tint)
        {
            SheetProperties.clrTabColor.SetThemeColor(TabColor, Tint);
            SheetProperties.HasTabColor = SheetProperties.clrTabColor.Color.IsEmpty ? false : true;
        }

        /// <summary>
        ///     Set normal margins.
        /// </summary>
        public void SetNormalMargins()
        {
            fTopMargin = SLConstants.NormalTopMargin;
            fBottomMargin = SLConstants.NormalBottomMargin;
            fLeftMargin = SLConstants.NormalLeftMargin;
            fRightMargin = SLConstants.NormalRightMargin;
            fHeaderMargin = SLConstants.NormalHeaderMargin;
            fFooterMargin = SLConstants.NormalFooterMargin;
            HasPageMargins = true;
        }

        /// <summary>
        ///     Set wide margins.
        /// </summary>
        public void SetWideMargins()
        {
            fTopMargin = SLConstants.WideTopMargin;
            fBottomMargin = SLConstants.WideBottomMargin;
            fLeftMargin = SLConstants.WideLeftMargin;
            fRightMargin = SLConstants.WideRightMargin;
            fHeaderMargin = SLConstants.WideHeaderMargin;
            fFooterMargin = SLConstants.WideFooterMargin;
            HasPageMargins = true;
        }

        /// <summary>
        ///     Set narrow margins.
        /// </summary>
        public void SetNarrowMargins()
        {
            fTopMargin = SLConstants.NarrowTopMargin;
            fBottomMargin = SLConstants.NarrowBottomMargin;
            fLeftMargin = SLConstants.NarrowLeftMargin;
            fRightMargin = SLConstants.NarrowRightMargin;
            fHeaderMargin = SLConstants.NarrowHeaderMargin;
            fFooterMargin = SLConstants.NarrowFooterMargin;
            HasPageMargins = true;
        }

        /// <summary>
        ///     Adjust the page a given percentage of the normal size.
        /// </summary>
        /// <param name="ScalePercentage">The scale percentage between 10% and 400%.</param>
        public void ScalePage(uint ScalePercentage)
        {
            if (ScalePercentage < 10) ScalePercentage = 10;
            if (ScalePercentage > 400) ScalePercentage = 400;
            iScale = ScalePercentage;

            iFitToWidth = 1;
            iFitToHeight = 1;

            SheetProperties.FitToPage = false;
        }

        /// <summary>
        ///     Fit to a given number of pages wide, and a given number of pages high.
        /// </summary>
        /// <param name="FitToWidth">Number of pages wide. Minimum is 1 page (default).</param>
        /// <param name="FitToHeight">Number of pages high. Minimum is 1 page (default).</param>
        public void ScalePage(uint FitToWidth, uint FitToHeight)
        {
            if (FitToWidth < 1) FitToWidth = 1;
            if (FitToHeight < 1) FitToHeight = 1;
            iFitToWidth = FitToWidth;
            iFitToHeight = FitToHeight;

            iScale = 100;

            SheetProperties.FitToPage = true;
        }

        // the switches are universal!
        // To do it "properly", we could have 6 versions of the switch variables
        // (1 for each type of odd/even/first header/footer).
        // Is this important? The "workaround" is to do each header/footer type
        // all the way through before working on another type. (which is more natural...)
        // The "bug" will appear if you append some styled text on say OddHeader,
        // then append some styled text on FirstFooter.
        // The switches are still assumed to work on OddHeader, but should be reset for
        // FirstFooter. This is fine until you go back to appending some styled text for
        // OddHeader.

        /// <summary>
        ///     Append text to the odd-numbered page header.
        /// </summary>
        /// <param name="Text">The text to be appended.</param>
        public void AppendOddHeader(string Text)
        {
            if (HFSection != SLHeaderFooterSection.OddHeader) ResetSwitches();
            OddHeaderText += Text;
            HFSection = SLHeaderFooterSection.OddHeader;
        }

        /// <summary>
        ///     Append text to the odd-numbered page header.
        /// </summary>
        /// <param name="FontStyle">The font style of the text.</param>
        /// <param name="Text">The text to be appended.</param>
        public void AppendOddHeader(SLFont FontStyle, string Text)
        {
            if (HFSection != SLHeaderFooterSection.OddHeader) ResetSwitches();
            OddHeaderText += string.Format("{0} {1}", StyleToAppend(FontStyle), Text);
            HFSection = SLHeaderFooterSection.OddHeader;
        }

        /// <summary>
        ///     Append a format code to the odd-numbered page header.
        /// </summary>
        /// <param name="Code">The format code.</param>
        public void AppendOddHeader(SLHeaderFooterFormatCodeValues Code)
        {
            if (HFSection != SLHeaderFooterSection.OddHeader) ResetSwitches();
            OddHeaderText += TextToAppend(Code);
            HFSection = SLHeaderFooterSection.OddHeader;
        }

        /// <summary>
        ///     Append text to the odd-numbered page footer.
        /// </summary>
        /// <param name="Text">The text to be appended.</param>
        public void AppendOddFooter(string Text)
        {
            if (HFSection != SLHeaderFooterSection.OddFooter) ResetSwitches();
            OddFooterText += Text;
            HFSection = SLHeaderFooterSection.OddFooter;
        }

        /// <summary>
        ///     Append text to the odd-numbered page footer.
        /// </summary>
        /// <param name="FontStyle">The font style of the text.</param>
        /// <param name="Text">The text to be appended.</param>
        public void AppendOddFooter(SLFont FontStyle, string Text)
        {
            if (HFSection != SLHeaderFooterSection.OddFooter) ResetSwitches();
            OddFooterText += string.Format("{0} {1}", StyleToAppend(FontStyle), Text);
            HFSection = SLHeaderFooterSection.OddFooter;
        }

        /// <summary>
        ///     Append a format code to the odd-numbered page footer.
        /// </summary>
        /// <param name="Code">The format code.</param>
        public void AppendOddFooter(SLHeaderFooterFormatCodeValues Code)
        {
            if (HFSection != SLHeaderFooterSection.OddFooter) ResetSwitches();
            OddFooterText += TextToAppend(Code);
            HFSection = SLHeaderFooterSection.OddFooter;
        }

        /// <summary>
        ///     Append text to the even-numbered page header.
        /// </summary>
        /// <param name="Text">The text to be appended.</param>
        public void AppendEvenHeader(string Text)
        {
            if (HFSection != SLHeaderFooterSection.EvenHeader) ResetSwitches();
            EvenHeaderText += Text;
            HFSection = SLHeaderFooterSection.EvenHeader;
        }

        /// <summary>
        ///     Append text to the even-numbered page header.
        /// </summary>
        /// <param name="FontStyle">The font style of the text.</param>
        /// <param name="Text">The text to be appended.</param>
        public void AppendEvenHeader(SLFont FontStyle, string Text)
        {
            if (HFSection != SLHeaderFooterSection.EvenHeader) ResetSwitches();
            EvenHeaderText += string.Format("{0} {1}", StyleToAppend(FontStyle), Text);
            HFSection = SLHeaderFooterSection.EvenHeader;
        }

        /// <summary>
        ///     Append a format code to the even-numbered page header.
        /// </summary>
        /// <param name="Code">The format code.</param>
        public void AppendEvenHeader(SLHeaderFooterFormatCodeValues Code)
        {
            if (HFSection != SLHeaderFooterSection.EvenHeader) ResetSwitches();
            EvenHeaderText += TextToAppend(Code);
            HFSection = SLHeaderFooterSection.EvenHeader;
        }

        /// <summary>
        ///     Append text to the even-numbered page footer.
        /// </summary>
        /// <param name="Text">The text to be appended.</param>
        public void AppendEvenFooter(string Text)
        {
            if (HFSection != SLHeaderFooterSection.EvenFooter) ResetSwitches();
            EvenFooterText += Text;
            HFSection = SLHeaderFooterSection.EvenFooter;
        }

        /// <summary>
        ///     Append text to the even-numbered page footer.
        /// </summary>
        /// <param name="FontStyle">The font style of the text.</param>
        /// <param name="Text">The text to be appended.</param>
        public void AppendEvenFooter(SLFont FontStyle, string Text)
        {
            if (HFSection != SLHeaderFooterSection.EvenFooter) ResetSwitches();
            EvenFooterText += string.Format("{0} {1}", StyleToAppend(FontStyle), Text);
            HFSection = SLHeaderFooterSection.EvenFooter;
        }

        /// <summary>
        ///     Append a format code to the even-numbered page footer.
        /// </summary>
        /// <param name="Code">The format code.</param>
        public void AppendEvenFooter(SLHeaderFooterFormatCodeValues Code)
        {
            if (HFSection != SLHeaderFooterSection.EvenFooter) ResetSwitches();
            EvenFooterText += TextToAppend(Code);
            HFSection = SLHeaderFooterSection.EvenFooter;
        }

        /// <summary>
        ///     Append text to the first page header.
        /// </summary>
        /// <param name="Text">The text to be appended.</param>
        public void AppendFirstHeader(string Text)
        {
            if (HFSection != SLHeaderFooterSection.FirstHeader) ResetSwitches();
            FirstHeaderText += Text;
            HFSection = SLHeaderFooterSection.FirstHeader;
        }

        /// <summary>
        ///     Append text to the first page header.
        /// </summary>
        /// <param name="FontStyle">The font style of the text.</param>
        /// <param name="Text">The text to be appended.</param>
        public void AppendFirstHeader(SLFont FontStyle, string Text)
        {
            if (HFSection != SLHeaderFooterSection.FirstHeader) ResetSwitches();
            FirstHeaderText += string.Format("{0} {1}", StyleToAppend(FontStyle), Text);
            HFSection = SLHeaderFooterSection.FirstHeader;
        }

        /// <summary>
        ///     Append a format code to the first page header.
        /// </summary>
        /// <param name="Code">The format code.</param>
        public void AppendFirstHeader(SLHeaderFooterFormatCodeValues Code)
        {
            if (HFSection != SLHeaderFooterSection.FirstHeader) ResetSwitches();
            FirstHeaderText += TextToAppend(Code);
            HFSection = SLHeaderFooterSection.FirstHeader;
        }

        /// <summary>
        ///     Append text to the first page footer.
        /// </summary>
        /// <param name="Text">The text to be appended.</param>
        public void AppendFirstFooter(string Text)
        {
            if (HFSection != SLHeaderFooterSection.FirstFooter) ResetSwitches();
            FirstFooterText += Text;
            HFSection = SLHeaderFooterSection.FirstFooter;
        }

        /// <summary>
        ///     Append text to the first page footer.
        /// </summary>
        /// <param name="FontStyle">The font style of the text.</param>
        /// <param name="Text">The text to be appended.</param>
        public void AppendFirstFooter(SLFont FontStyle, string Text)
        {
            if (HFSection != SLHeaderFooterSection.FirstFooter) ResetSwitches();
            FirstFooterText += string.Format("{0} {1}", StyleToAppend(FontStyle), Text);
            HFSection = SLHeaderFooterSection.FirstFooter;
        }

        /// <summary>
        ///     Append a format code to the first page footer.
        /// </summary>
        /// <param name="Code">The format code.</param>
        public void AppendFirstFooter(SLHeaderFooterFormatCodeValues Code)
        {
            if (HFSection != SLHeaderFooterSection.FirstFooter) ResetSwitches();
            FirstFooterText += TextToAppend(Code);
            HFSection = SLHeaderFooterSection.FirstFooter;
        }

        /// <summary>
        ///     Get the text from the left section of the header.
        /// </summary>
        /// <returns>The text.</returns>
        public string GetLeftHeaderText()
        {
            return GetHeaderFooterText(true, SLHeaderFooterTypeValues.Odd, SLHeaderFooterSectionValues.Left);
        }

        /// <summary>
        ///     Get the text from the left section of the header.
        /// </summary>
        /// <param name="HeaderType">The header type.</param>
        /// <returns>The text.</returns>
        public string GetLeftHeaderText(SLHeaderFooterTypeValues HeaderType)
        {
            return GetHeaderFooterText(true, HeaderType, SLHeaderFooterSectionValues.Left);
        }

        /// <summary>
        ///     Get the text from the center section of the header.
        /// </summary>
        /// <returns>The text.</returns>
        public string GetCenterHeaderText()
        {
            return GetHeaderFooterText(true, SLHeaderFooterTypeValues.Odd, SLHeaderFooterSectionValues.Center);
        }

        /// <summary>
        ///     Get the text from the center section of the header.
        /// </summary>
        /// <param name="HeaderType">The header type.</param>
        /// <returns>The text.</returns>
        public string GetCenterHeaderText(SLHeaderFooterTypeValues HeaderType)
        {
            return GetHeaderFooterText(true, HeaderType, SLHeaderFooterSectionValues.Center);
        }

        /// <summary>
        ///     Get the text from the right section of the header.
        /// </summary>
        /// <returns>The text.</returns>
        public string GetRightHeaderText()
        {
            return GetHeaderFooterText(true, SLHeaderFooterTypeValues.Odd, SLHeaderFooterSectionValues.Right);
        }

        /// <summary>
        ///     Get the text from the right section of the header.
        /// </summary>
        /// <param name="HeaderType">The header type.</param>
        /// <returns>The text.</returns>
        public string GetRightHeaderText(SLHeaderFooterTypeValues HeaderType)
        {
            return GetHeaderFooterText(true, HeaderType, SLHeaderFooterSectionValues.Right);
        }

        /// <summary>
        ///     Get the text from the left section of the footer.
        /// </summary>
        /// <returns>The text.</returns>
        public string GetLeftFooterText()
        {
            return GetHeaderFooterText(false, SLHeaderFooterTypeValues.Odd, SLHeaderFooterSectionValues.Left);
        }

        /// <summary>
        ///     Get the text from the left section of the footer.
        /// </summary>
        /// <param name="FooterType">The footer type.</param>
        /// <returns>The text.</returns>
        public string GetLeftFooterText(SLHeaderFooterTypeValues FooterType)
        {
            return GetHeaderFooterText(false, FooterType, SLHeaderFooterSectionValues.Left);
        }

        /// <summary>
        ///     Get the text from the center section of the footer.
        /// </summary>
        /// <returns>The text.</returns>
        public string GetCenterFooterText()
        {
            return GetHeaderFooterText(false, SLHeaderFooterTypeValues.Odd, SLHeaderFooterSectionValues.Center);
        }

        /// <summary>
        ///     Get the text from the center section of the footer.
        /// </summary>
        /// <param name="FooterType">The footer type.</param>
        /// <returns>The text.</returns>
        public string GetCenterFooterText(SLHeaderFooterTypeValues FooterType)
        {
            return GetHeaderFooterText(false, FooterType, SLHeaderFooterSectionValues.Center);
        }

        /// <summary>
        ///     Get the text from the right section of the footer.
        /// </summary>
        /// <returns>The text.</returns>
        public string GetRightFooterText()
        {
            return GetHeaderFooterText(false, SLHeaderFooterTypeValues.Odd, SLHeaderFooterSectionValues.Right);
        }

        /// <summary>
        ///     Get the text from the right section of the footer.
        /// </summary>
        /// <param name="FooterType">The footer type.</param>
        /// <returns>The text.</returns>
        public string GetRightFooterText(SLHeaderFooterTypeValues FooterType)
        {
            return GetHeaderFooterText(false, FooterType, SLHeaderFooterSectionValues.Right);
        }

        /// <summary>
        ///     Set the text of the left section of the header.
        /// </summary>
        /// <param name="Text">The text.</param>
        public void SetLeftHeaderText(string Text)
        {
            SetHeaderFooterText(true, SLHeaderFooterTypeValues.Odd, SLHeaderFooterSectionValues.Left, Text);
        }

        /// <summary>
        ///     Set the text of the left section of the header.
        /// </summary>
        /// <param name="HeaderType">The header type.</param>
        /// <param name="Text">The text.</param>
        public void SetLeftHeaderText(SLHeaderFooterTypeValues HeaderType, string Text)
        {
            SetHeaderFooterText(true, HeaderType, SLHeaderFooterSectionValues.Left, Text);
        }

        /// <summary>
        ///     Set the text of the center section of the header.
        /// </summary>
        /// <param name="Text">The text.</param>
        public void SetCenterHeaderText(string Text)
        {
            SetHeaderFooterText(true, SLHeaderFooterTypeValues.Odd, SLHeaderFooterSectionValues.Center, Text);
        }

        /// <summary>
        ///     Set the text of the center section of the header.
        /// </summary>
        /// <param name="HeaderType">The header type.</param>
        /// <param name="Text">The text.</param>
        public void SetCenterHeaderText(SLHeaderFooterTypeValues HeaderType, string Text)
        {
            SetHeaderFooterText(true, HeaderType, SLHeaderFooterSectionValues.Center, Text);
        }

        /// <summary>
        ///     Set the text of the right section of the header.
        /// </summary>
        /// <param name="Text">The text.</param>
        public void SetRightHeaderText(string Text)
        {
            SetHeaderFooterText(true, SLHeaderFooterTypeValues.Odd, SLHeaderFooterSectionValues.Right, Text);
        }

        /// <summary>
        ///     Set the text of the right section of the header.
        /// </summary>
        /// <param name="HeaderType">The header type.</param>
        /// <param name="Text">The text.</param>
        public void SetRightHeaderText(SLHeaderFooterTypeValues HeaderType, string Text)
        {
            SetHeaderFooterText(true, HeaderType, SLHeaderFooterSectionValues.Right, Text);
        }

        /// <summary>
        ///     Set the text of the left section of the footer.
        /// </summary>
        /// <param name="Text">The text.</param>
        public void SetLeftFooterText(string Text)
        {
            SetHeaderFooterText(false, SLHeaderFooterTypeValues.Odd, SLHeaderFooterSectionValues.Left, Text);
        }

        /// <summary>
        ///     Set the text of the left section of the footer.
        /// </summary>
        /// <param name="FooterType">The footer type.</param>
        /// <param name="Text">The text.</param>
        public void SetLeftFooterText(SLHeaderFooterTypeValues FooterType, string Text)
        {
            SetHeaderFooterText(false, FooterType, SLHeaderFooterSectionValues.Left, Text);
        }

        /// <summary>
        ///     Set the text of the center section of the footer.
        /// </summary>
        /// <param name="Text">The text.</param>
        public void SetCenterFooterText(string Text)
        {
            SetHeaderFooterText(false, SLHeaderFooterTypeValues.Odd, SLHeaderFooterSectionValues.Center, Text);
        }

        /// <summary>
        ///     Set the text of the center section of the footer.
        /// </summary>
        /// <param name="FooterType">The footer type.</param>
        /// <param name="Text">The text.</param>
        public void SetCenterFooterText(SLHeaderFooterTypeValues FooterType, string Text)
        {
            SetHeaderFooterText(false, FooterType, SLHeaderFooterSectionValues.Center, Text);
        }

        /// <summary>
        ///     Set the text of the right section of the footer.
        /// </summary>
        /// <param name="Text">The text.</param>
        public void SetRightFooterText(string Text)
        {
            SetHeaderFooterText(false, SLHeaderFooterTypeValues.Odd, SLHeaderFooterSectionValues.Right, Text);
        }

        /// <summary>
        ///     Set the text of the right section of the footer.
        /// </summary>
        /// <param name="FooterType">The footer type.</param>
        /// <param name="Text">The text.</param>
        public void SetRightFooterText(SLHeaderFooterTypeValues FooterType, string Text)
        {
            SetHeaderFooterText(false, FooterType, SLHeaderFooterSectionValues.Right, Text);
        }

        private string GetHeaderFooterText(bool IsHeader, SLHeaderFooterTypeValues Type,
            SLHeaderFooterSectionValues Section)
        {
            var result = string.Empty;
            string sLeft = string.Empty, sCenter = string.Empty, sRight = string.Empty;
            if (IsHeader)
            {
                if (Type == SLHeaderFooterTypeValues.Even)
                    SplitHeaderFooterText(EvenHeaderText, out sLeft, out sCenter, out sRight);
                else if (Type == SLHeaderFooterTypeValues.First)
                    SplitHeaderFooterText(FirstHeaderText, out sLeft, out sCenter, out sRight);
                else SplitHeaderFooterText(OddHeaderText, out sLeft, out sCenter, out sRight);
            }
            else
            {
                if (Type == SLHeaderFooterTypeValues.Even)
                    SplitHeaderFooterText(EvenFooterText, out sLeft, out sCenter, out sRight);
                else if (Type == SLHeaderFooterTypeValues.First)
                    SplitHeaderFooterText(FirstFooterText, out sLeft, out sCenter, out sRight);
                else SplitHeaderFooterText(OddFooterText, out sLeft, out sCenter, out sRight);
            }

            if (Section == SLHeaderFooterSectionValues.Left) result = sLeft;
            else if (Section == SLHeaderFooterSectionValues.Right) result = sRight;
            else result = sCenter;

            result = TranslateToUserFriendlyCode(result);

            return result;
        }

        private void SetHeaderFooterText(bool IsHeader, SLHeaderFooterTypeValues Type,
            SLHeaderFooterSectionValues Section, string Text)
        {
            var result = TranslateToInternalCode(Text);
            string sLeft = string.Empty, sCenter = string.Empty, sRight = string.Empty;
            if (IsHeader)
            {
                if (Type == SLHeaderFooterTypeValues.Even)
                    SplitHeaderFooterText(EvenHeaderText, out sLeft, out sCenter, out sRight);
                else if (Type == SLHeaderFooterTypeValues.First)
                    SplitHeaderFooterText(FirstHeaderText, out sLeft, out sCenter, out sRight);
                else SplitHeaderFooterText(OddHeaderText, out sLeft, out sCenter, out sRight);
            }
            else
            {
                if (Type == SLHeaderFooterTypeValues.Even)
                    SplitHeaderFooterText(EvenFooterText, out sLeft, out sCenter, out sRight);
                else if (Type == SLHeaderFooterTypeValues.First)
                    SplitHeaderFooterText(FirstFooterText, out sLeft, out sCenter, out sRight);
                else SplitHeaderFooterText(OddFooterText, out sLeft, out sCenter, out sRight);
            }

            if (Section == SLHeaderFooterSectionValues.Left) sLeft = result;
            else if (Section == SLHeaderFooterSectionValues.Right) sRight = result;
            else sCenter = result;

            result = string.Empty;

            if (sLeft.Length > 0) result += "&L" + sLeft;
            if (sCenter.Length > 0) result += "&C" + sCenter;
            if (sRight.Length > 0) result += "&R" + sRight;

            if (IsHeader)
            {
                if (Type == SLHeaderFooterTypeValues.Even) EvenHeaderText = result;
                else if (Type == SLHeaderFooterTypeValues.First) FirstHeaderText = result;
                else OddHeaderText = result;
            }
            else
            {
                if (Type == SLHeaderFooterTypeValues.Even) EvenFooterText = result;
                else if (Type == SLHeaderFooterTypeValues.First) FirstFooterText = result;
                else OddFooterText = result;
            }
        }

        private string TranslateToUserFriendlyCode(string HeaderFooterText)
        {
            var result = HeaderFooterText;
            result = Regex.Replace(result, "&[Pp]", "&[Page]");
            result = Regex.Replace(result, "&[Nn]", "&[Pages]");
            result = Regex.Replace(result, "&[Dd]", "&[Date]");
            result = Regex.Replace(result, "&[Tt]", "&[Time]");
            result = Regex.Replace(result, "&[Zz]", "&[Path]");
            result = Regex.Replace(result, "&[Ff]", "&[File]");
            result = Regex.Replace(result, "&[Aa]", "&[Tab]");

            return result;
        }

        private string TranslateToInternalCode(string HeaderFooterText)
        {
            var result = HeaderFooterText;
            result = Regex.Replace(result, "&\\[Page\\]", "&P");
            result = Regex.Replace(result, "&\\[Pages\\]", "&N");
            result = Regex.Replace(result, "&\\[Date\\]", "&D");
            result = Regex.Replace(result, "&\\[Time\\]", "&T");
            result = Regex.Replace(result, "&\\[Path\\]", "&Z");
            result = Regex.Replace(result, "&\\[File\\]", "&F");
            result = Regex.Replace(result, "&\\[Tab\\]", "&A");

            return result;
        }

        private string StyleToAppend(SLFont ft)
        {
            var result = string.Empty;

            var sBoldItalic = string.Empty;
            if ((ft.Bold != null) && ft.Bold.Value && (ft.Italic != null) && ft.Italic.Value)
                sBoldItalic = "Bold Italic";
            else if ((ft.Bold != null) && ft.Bold.Value)
                sBoldItalic = "Bold";
            else if ((ft.Italic != null) && ft.Italic.Value)
                sBoldItalic = "Italic";
            else
                sBoldItalic = "Regular";

            // if it's bold or italic, there must at least be a font name or font scheme
            if (!sBoldItalic.Equals("Regular"))
                if ((ft.FontName == null) || (ft.FontName.Length == 0))
                    if (ft.FontScheme == FontSchemeValues.None) ft.FontScheme = FontSchemeValues.Minor;

            var sFontStyle = string.Empty;
            if (FontStyleSwitch)
            {
                sFontStyle = FontStyleToAppend(ft, sBoldItalic);
                // must write something
                if (sFontStyle.Length == 0) sFontStyle = "&\"-,Regular\"";
            }
            else
            {
                sFontStyle = FontStyleToAppend(ft, sBoldItalic);
            }

            if ((sFontStyle.Length == 0) || sFontStyle.Equals("&\"-,Regular\""))
                FontStyleSwitch = false;
            else
                FontStyleSwitch = true;

            result += sFontStyle;

            if (FontSizeSwitch)
            {
                // font size switch is on, so must write something
                if (ft.FontSize != null)
                    result += string.Format("&{0}", (int) ft.FontSize);
                else
                    result += string.Format("&{0}", (int) SLConstants.DefaultFontSize);
                FontSizeSwitch = false;
            }
            else
            {
                if ((ft.FontSize != null) && ((int) ft.FontSize != (int) SLConstants.DefaultFontSize))
                {
                    result += string.Format("&{0}", (int) ft.FontSize);
                    FontSizeSwitch = true;
                }
            }

            if (StrikeSwitch)
            {
                // already in strikethrough mode
                // so only write something if given font style has no strikethrough
                if ((ft.Strike == null) || ((ft.Strike != null) && !ft.Strike.Value))
                {
                    result += "&S";
                    StrikeSwitch = false;
                }
            }
            else
            {
                if ((ft.Strike != null) && ft.Strike.Value)
                {
                    result += "&S";
                    StrikeSwitch = true;
                }
            }

            if (SuperscriptSwitch)
            {
                // already in superscript mode
                // so only write something if given font style has no superscript
                if (!ft.HasVerticalAlignment
                    || (ft.HasVerticalAlignment && (ft.VerticalAlignment != VerticalAlignmentRunValues.Superscript)))
                {
                    result += "&X";
                    SuperscriptSwitch = false;
                }
            }
            else
            {
                if (ft.HasVerticalAlignment && (ft.VerticalAlignment == VerticalAlignmentRunValues.Superscript))
                {
                    result += "&X";
                    SuperscriptSwitch = true;
                }
            }

            if (SubscriptSwitch)
            {
                // already in subscript mode
                // so only write something if given font style has no subscript
                if (!ft.HasVerticalAlignment
                    || (ft.HasVerticalAlignment && (ft.VerticalAlignment != VerticalAlignmentRunValues.Subscript)))
                {
                    result += "&Y";
                    SubscriptSwitch = false;
                }
            }
            else
            {
                if (ft.HasVerticalAlignment && (ft.VerticalAlignment == VerticalAlignmentRunValues.Subscript))
                {
                    result += "&Y";
                    SubscriptSwitch = true;
                }
            }

            if (UnderlineSwitch)
            {
                // already in underline mode
                // so only write something if given font style has no underline
                if (!ft.HasUnderline
                    || (ft.HasUnderline && (ft.Underline != UnderlineValues.Single)))
                {
                    // take care of SingleAccounting?
                    result += "&U";
                    UnderlineSwitch = false;
                }
            }
            else
            {
                if (ft.HasUnderline && (ft.Underline == UnderlineValues.Single))
                {
                    // take care of SingleAccounting?
                    result += "&U";
                    UnderlineSwitch = true;
                }
            }

            if (DoubleUnderlineSwitch)
            {
                // already in double underline mode
                // so only write something if given font style has no double underline
                if (!ft.HasUnderline
                    || (ft.HasUnderline && (ft.Underline == UnderlineValues.Double)))
                {
                    // take care of DoubleAccounting?
                    result += "&E";
                    DoubleUnderlineSwitch = false;
                }
            }
            else
            {
                if (ft.HasUnderline && (ft.Underline == UnderlineValues.Double))
                {
                    // take care of DoubleAccounting?
                    result += "&E";
                    DoubleUnderlineSwitch = true;
                }
            }

            if (FontColorSwitch)
            {
                if (ft.HasFontColor)
                {
                    result += FontColorToAppend(ft.clrFontColor);
                }
                else
                {
                    result += "&K01+000";
                    FontColorSwitch = false;
                }
            }
            else
            {
                if (ft.HasFontColor)
                {
                    result += FontColorToAppend(ft.clrFontColor);
                    FontColorSwitch = true;
                }
            }

            return result;
        }

        private string FontStyleToAppend(SLFont ft, string BoldItalic)
        {
            var result = string.Empty;

            if (ft.HasFontScheme)
                if (ft.FontScheme == FontSchemeValues.Minor)
                {
                    result = string.Format("&\"-,{0}\"", BoldItalic);
                }
                else if (ft.FontScheme == FontSchemeValues.Major)
                {
                    result = string.Format("&\"+,{0}\"", BoldItalic);
                }
                else
                {
                    if ((ft.FontName != null) && (ft.FontName.Length > 0))
                        result = string.Format("&\"{1},{0}\"", BoldItalic, ft.FontName);
                    else
                        result = string.Format("&\"-,{0}\"", BoldItalic);
                }
            else if ((ft.FontName != null) && (ft.FontName.Length > 0))
                result = string.Format("&\"{1},{0}\"", BoldItalic, ft.FontName);

            return result;
        }

        private string FontColorToAppend(SLColor clr)
        {
            var result = "&K01+000";

            if (clr.Theme != null)
            {
                var fTint = 0.0;
                var bPositive = true;
                var sTint = string.Empty;
                if (clr.Tint != null)
                    fTint = clr.Tint.Value;

                if (fTint < 0)
                {
                    fTint = -fTint;
                    bPositive = false;
                }
                sTint = fTint.ToString(CultureInfo.InvariantCulture).Replace(".", "").PadRight(3, '0').Substring(0, 3);

                result = string.Format("&K{0}{1}{2}", clr.Theme.Value.ToString("d2"), bPositive ? "+" : "-", sTint);
            }
            else
            {
                result = string.Format("&K{0}{1}{2}", clr.Color.R.ToString("X2"), clr.Color.G.ToString("X2"),
                    clr.Color.B.ToString("X2"));
            }

            return result;
        }

        private string TextToAppend(SLHeaderFooterFormatCodeValues Code)
        {
            var result = string.Empty;
            switch (Code)
            {
                case SLHeaderFooterFormatCodeValues.Left:
                    result = "&L";
                    ResetSwitches();
                    break;
                case SLHeaderFooterFormatCodeValues.Center:
                    result = "&C";
                    ResetSwitches();
                    break;
                case SLHeaderFooterFormatCodeValues.Right:
                    result = "&R";
                    ResetSwitches();
                    break;
                case SLHeaderFooterFormatCodeValues.PageNumber:
                    result = "&P";
                    break;
                case SLHeaderFooterFormatCodeValues.NumberOfPages:
                    result = "&N";
                    break;
                case SLHeaderFooterFormatCodeValues.Date:
                    result = "&D";
                    break;
                case SLHeaderFooterFormatCodeValues.Time:
                    result = "&T";
                    break;
                case SLHeaderFooterFormatCodeValues.FilePath:
                    result = "&Z";
                    break;
                case SLHeaderFooterFormatCodeValues.FileName:
                    result = "&F";
                    break;
                case SLHeaderFooterFormatCodeValues.SheetName:
                    result = "&A";
                    break;
                case SLHeaderFooterFormatCodeValues.ResetFont:
                    if (FontStyleSwitch) result += "&\"-,Regular\"";
                    if (FontSizeSwitch) result += string.Format("&{0}", (int) SLConstants.DefaultFontSize);
                    if (StrikeSwitch) result += "&S";
                    if (SuperscriptSwitch) result += "&X";
                    if (SubscriptSwitch) result += "&Y";
                    if (UnderlineSwitch) result += "&U";
                    if (DoubleUnderlineSwitch) result += "&E";
                    if (FontColorSwitch) result += "&K01+000";
                    ResetSwitches();
                    break;
            }
            return result;
        }

        private void SplitHeaderFooterText(string Text, out string Left, out string Center, out string Right)
        {
            Left = string.Empty;
            Center = string.Empty;
            Right = string.Empty;

            var sbLeft = new StringBuilder();
            var sbCenter = new StringBuilder();
            var sbRight = new StringBuilder();

            // 0-left, 1-center, 2-right
            var iChoice = 1;

            for (var i = 0; i < Text.Length; ++i)
                if (Text[i] == '&')
                {
                    if (i + 1 < Text.Length)
                    {
                        // still within string length
                        if ((Text[i + 1] == 'L') || (Text[i + 1] == 'l'))
                        {
                            iChoice = 0;
                            ++i;
                        }
                        else if ((Text[i + 1] == 'C') || (Text[i + 1] == 'c'))
                        {
                            iChoice = 1;
                            ++i;
                        }
                        else if ((Text[i + 1] == 'R') || (Text[i + 1] == 'r'))
                        {
                            iChoice = 2;
                            ++i;
                        }
                        else
                        {
                            // we're appending basically the ampersand
                            if (iChoice == 0) sbLeft.Append(Text[i]);
                            else if (iChoice == 2) sbRight.Append(Text[i]);
                            else sbCenter.Append(Text[i]);
                        }
                    }
                    else
                    {
                        // we're appending basically the ampersand
                        if (iChoice == 0) sbLeft.Append(Text[i]);
                        else if (iChoice == 2) sbRight.Append(Text[i]);
                        else sbCenter.Append(Text[i]);
                    }
                }
                else
                {
                    if (iChoice == 0) sbLeft.Append(Text[i]);
                    else if (iChoice == 2) sbRight.Append(Text[i]);
                    else sbCenter.Append(Text[i]);
                }

            Left = sbLeft.ToString();
            Center = sbCenter.ToString();
            Right = sbRight.ToString();
        }

        internal void ImportPrintOptions(PrintOptions po)
        {
            if (po.HorizontalCentered != null) PrintHorizontalCentered = po.HorizontalCentered.Value;
            if (po.VerticalCentered != null) PrintVerticalCentered = po.VerticalCentered.Value;
            if (po.Headings != null) PrintHeadings = po.Headings.Value;
            if (po.GridLines != null) PrintGridLines = po.GridLines.Value;
            if (po.GridLinesSet != null) PrintGridLinesSet = po.GridLinesSet.Value;
        }

        internal PrintOptions ExportPrintOptions()
        {
            var po = new PrintOptions();
            if (PrintHorizontalCentered) po.HorizontalCentered = true;
            if (PrintVerticalCentered) po.VerticalCentered = true;
            if (PrintHeadings) po.Headings = true;
            if (PrintGridLines) po.GridLines = true;
            if (!PrintGridLinesSet) po.GridLinesSet = false;

            return po;
        }

        internal void ImportPageMargins(PageMargins pm)
        {
            if (pm.Left != null) LeftMargin = pm.Left.Value;
            if (pm.Right != null) RightMargin = pm.Right.Value;
            if (pm.Top != null) TopMargin = pm.Top.Value;
            if (pm.Bottom != null) BottomMargin = pm.Bottom.Value;
            if (pm.Header != null) HeaderMargin = pm.Header.Value;
            if (pm.Footer != null) FooterMargin = pm.Footer.Value;
        }

        internal PageMargins ExportPageMargins()
        {
            var pm = new PageMargins();
            pm.Left = LeftMargin;
            pm.Right = RightMargin;
            pm.Top = TopMargin;
            pm.Bottom = BottomMargin;
            pm.Header = HeaderMargin;
            pm.Footer = FooterMargin;

            return pm;
        }

        internal void ImportPageSetup(PageSetup ps)
        {
            if (ps.PaperSize != null)
                if (Enum.IsDefined(typeof(SLPaperSizeValues), ps.PaperSize.Value))
                    PaperSize = (SLPaperSizeValues) ps.PaperSize.Value;
                else
                    PaperSize = SLPaperSizeValues.LetterPaper;

            if (ps.Scale != null) iScale = ps.Scale.Value;
            if (ps.FirstPageNumber != null) iFirstPageNumber = ps.FirstPageNumber.Value;
            if (ps.FitToWidth != null) iFitToWidth = ps.FitToWidth.Value;
            if (ps.FitToHeight != null) iFitToHeight = ps.FitToHeight.Value;
            if (ps.PageOrder != null) PageOrder = ps.PageOrder.Value;
            if (ps.Orientation != null) Orientation = ps.Orientation.Value;
            if (ps.UsePrinterDefaults != null) UsePrinterDefaults = ps.UsePrinterDefaults.Value;
            if (ps.BlackAndWhite != null) BlackAndWhite = ps.BlackAndWhite.Value;
            if (ps.Draft != null) Draft = ps.Draft.Value;
            if (ps.CellComments != null) CellComments = ps.CellComments.Value;
            if (ps.Errors != null) Errors = ps.Errors.Value;
            if (ps.HorizontalDpi != null) HorizontalDpi = ps.HorizontalDpi.Value;
            if (ps.VerticalDpi != null) VerticalDpi = ps.VerticalDpi.Value;
            if (ps.Copies != null) Copies = ps.Copies.Value;
        }

        internal PageSetup ExportPageSetup()
        {
            var ps = new PageSetup();
            if (PaperSize != SLPaperSizeValues.LetterPaper) ps.PaperSize = (uint) PaperSize;
            if (Scale != 100) ps.Scale = Scale;
            if ((FitToWidth != 1) || (FitToHeight != 1))
            {
                ps.FitToWidth = FitToWidth;
                ps.FitToHeight = FitToHeight;
            }
            if (FirstPageNumber != 1)
            {
                ps.FirstPageNumber = FirstPageNumber;
                ps.UseFirstPageNumber = true;
            }
            if (PageOrder != PageOrderValues.DownThenOver) ps.PageOrder = PageOrder;
            if (Orientation != OrientationValues.Default) ps.Orientation = Orientation;
            if (!UsePrinterDefaults) ps.UsePrinterDefaults = UsePrinterDefaults;
            if (BlackAndWhite) ps.BlackAndWhite = BlackAndWhite;
            if (Draft) ps.Draft = Draft;
            if (CellComments != CellCommentsValues.None) ps.CellComments = CellComments;
            if (Errors != PrintErrorValues.Displayed) ps.Errors = Errors;
            if (HorizontalDpi != 600) ps.HorizontalDpi = HorizontalDpi;
            if (VerticalDpi != 600) ps.VerticalDpi = VerticalDpi;
            if (Copies != 1) ps.Copies = Copies;

            return ps;
        }

        internal void ImportChartSheetPageSetup(ChartSheetPageSetup ps)
        {
            if (ps.PaperSize != null)
                if (Enum.IsDefined(typeof(SLPaperSizeValues), ps.PaperSize.Value))
                    PaperSize = (SLPaperSizeValues) ps.PaperSize.Value;
                else
                    PaperSize = SLPaperSizeValues.LetterPaper;

            if (ps.FirstPageNumber != null) iFirstPageNumber = ps.FirstPageNumber.Value;
            if (ps.Orientation != null) Orientation = ps.Orientation.Value;
            if (ps.UsePrinterDefaults != null) UsePrinterDefaults = ps.UsePrinterDefaults.Value;
            if (ps.BlackAndWhite != null) BlackAndWhite = ps.BlackAndWhite.Value;
            if (ps.Draft != null) Draft = ps.Draft.Value;
            if (ps.HorizontalDpi != null) HorizontalDpi = ps.HorizontalDpi.Value;
            if (ps.VerticalDpi != null) VerticalDpi = ps.VerticalDpi.Value;
            if (ps.Copies != null) Copies = ps.Copies.Value;
        }

        internal ChartSheetPageSetup ExportChartSheetPageSetup()
        {
            var ps = new ChartSheetPageSetup();
            if (PaperSize != SLPaperSizeValues.LetterPaper) ps.PaperSize = (uint) PaperSize;
            if (FirstPageNumber != 1)
            {
                ps.FirstPageNumber = FirstPageNumber;
                ps.UseFirstPageNumber = true;
            }
            if (Orientation != OrientationValues.Default) ps.Orientation = Orientation;
            if (!UsePrinterDefaults) ps.UsePrinterDefaults = UsePrinterDefaults;
            if (BlackAndWhite) ps.BlackAndWhite = BlackAndWhite;
            if (Draft) ps.Draft = Draft;
            if (HorizontalDpi != 600) ps.HorizontalDpi = HorizontalDpi;
            if (VerticalDpi != 600) ps.VerticalDpi = VerticalDpi;
            if (Copies != 1) ps.Copies = Copies;

            return ps;
        }

        internal void ImportHeaderFooter(HeaderFooter hf)
        {
            if (hf.OddHeader != null) OddHeaderText = hf.OddHeader.Text;
            if (hf.OddFooter != null) OddFooterText = hf.OddFooter.Text;
            if (hf.EvenHeader != null) EvenHeaderText = hf.EvenHeader.Text;
            if (hf.EvenFooter != null) EvenFooterText = hf.EvenFooter.Text;
            if (hf.FirstHeader != null) FirstHeaderText = hf.FirstHeader.Text;
            if (hf.FirstFooter != null) FirstFooterText = hf.FirstFooter.Text;
            if (hf.DifferentOddEven != null) DifferentOddEvenPages = hf.DifferentOddEven.Value;
            if (hf.DifferentFirst != null) DifferentFirstPage = hf.DifferentFirst.Value;
            if (hf.ScaleWithDoc != null) ScaleWithDocument = hf.ScaleWithDoc.Value;
            if (hf.AlignWithMargins != null) AlignWithMargins = hf.AlignWithMargins.Value;
        }

        internal HeaderFooter ExportHeaderFooter()
        {
            var hf = new HeaderFooter();
            if (OddHeaderText.Length > 0) hf.OddHeader = new OddHeader(OddHeaderText);
            if (OddFooterText.Length > 0) hf.OddFooter = new OddFooter(OddFooterText);
            if (EvenHeaderText.Length > 0) hf.EvenHeader = new EvenHeader(EvenHeaderText);
            if (EvenFooterText.Length > 0) hf.EvenFooter = new EvenFooter(EvenFooterText);
            if (FirstHeaderText.Length > 0) hf.FirstHeader = new FirstHeader(FirstHeaderText);
            if (FirstFooterText.Length > 0) hf.FirstFooter = new FirstFooter(FirstFooterText);
            if (DifferentOddEvenPages) hf.DifferentOddEven = DifferentOddEvenPages;
            if (DifferentFirstPage) hf.DifferentFirst = DifferentFirstPage;
            if (!ScaleWithDocument) hf.ScaleWithDoc = ScaleWithDocument;
            if (!AlignWithMargins) hf.AlignWithMargins = AlignWithMargins;

            return hf;
        }

        internal SLSheetView ExportSLSheetView()
        {
            var sv = new SLSheetView();
            if (bShowFormulas != null) sv.ShowFormulas = bShowFormulas.Value;
            if (bShowGridLines != null) sv.ShowGridLines = bShowGridLines.Value;
            if (bShowRowColumnHeaders != null) sv.ShowRowColHeaders = bShowRowColumnHeaders.Value;
            if (bShowRuler != null) sv.ShowRuler = bShowRuler.Value;
            if (vView != null) sv.View = vView.Value;
            if (iZoomScale != null) sv.ZoomScale = iZoomScale.Value;
            if (iZoomScaleNormal != null) sv.ZoomScaleNormal = iZoomScaleNormal.Value;
            if (iZoomScalePageLayoutView != null) sv.ZoomScalePageLayoutView = iZoomScalePageLayoutView.Value;

            return sv;
        }

        internal SLPageSettings Clone()
        {
            var ps = new SLPageSettings(listThemeColors, listIndexedColors);

            ps.SheetProperties = SheetProperties.Clone();

            ps.bShowFormulas = bShowFormulas;
            ps.bShowGridLines = bShowGridLines;
            ps.bShowRowColumnHeaders = bShowRowColumnHeaders;
            ps.bShowRuler = bShowRuler;
            ps.vView = vView;
            ps.iZoomScale = iZoomScale;
            ps.iZoomScaleNormal = iZoomScaleNormal;
            ps.iZoomScalePageLayoutView = iZoomScalePageLayoutView;

            ps.PrintHorizontalCentered = PrintHorizontalCentered;
            ps.PrintVerticalCentered = PrintVerticalCentered;
            ps.PrintHeadings = PrintHeadings;
            ps.PrintGridLines = PrintGridLines;
            ps.PrintGridLinesSet = PrintGridLinesSet;

            ps.HasPageMargins = HasPageMargins;
            ps.fLeftMargin = fLeftMargin;
            ps.fRightMargin = fRightMargin;
            ps.fTopMargin = fTopMargin;
            ps.fBottomMargin = fBottomMargin;
            ps.fHeaderMargin = fHeaderMargin;
            ps.fFooterMargin = fFooterMargin;

            ps.PaperSize = PaperSize;
            ps.iFirstPageNumber = iFirstPageNumber;
            ps.iScale = iScale;
            ps.iFitToWidth = iFitToWidth;
            ps.iFitToHeight = iFitToHeight;
            ps.PageOrder = PageOrder;
            ps.Orientation = Orientation;
            ps.UsePrinterDefaults = UsePrinterDefaults;
            ps.BlackAndWhite = BlackAndWhite;
            ps.Draft = Draft;
            ps.CellComments = CellComments;
            ps.Errors = Errors;
            ps.HorizontalDpi = HorizontalDpi;
            ps.VerticalDpi = VerticalDpi;
            ps.iCopies = iCopies;

            ps.OddHeaderText = OddHeaderText;
            ps.OddFooterText = OddFooterText;
            ps.EvenHeaderText = EvenHeaderText;
            ps.EvenFooterText = EvenFooterText;
            ps.FirstHeaderText = FirstHeaderText;
            ps.FirstFooterText = FirstFooterText;
            ps.DifferentOddEvenPages = DifferentOddEvenPages;
            ps.DifferentFirstPage = DifferentFirstPage;
            ps.ScaleWithDocument = ScaleWithDocument;
            ps.AlignWithMargins = AlignWithMargins;

            return ps;
        }
    }
}