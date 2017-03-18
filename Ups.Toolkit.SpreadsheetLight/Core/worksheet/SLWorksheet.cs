using System.Collections.Generic;
using System.Drawing;
using DocumentFormat.OpenXml.Packaging;
using Ups.Toolkit.SpreadsheetLight.Core.conditionalformatting;
using Ups.Toolkit.SpreadsheetLight.Core.Charts;
using Ups.Toolkit.SpreadsheetLight.Core.Drawing;
using Ups.Toolkit.SpreadsheetLight.Core.misc;
using Ups.Toolkit.SpreadsheetLight.Core.office2010;
using Ups.Toolkit.SpreadsheetLight.Core.sparkline;
using Ups.Toolkit.SpreadsheetLight.Core.table;

namespace Ups.Toolkit.SpreadsheetLight.Core.worksheet
{
    internal class SLWorksheet
    {
        internal bool HasAutoFilter;

        // note that this doesn't mean that the worksheet is protected,
        // just that the SheetProtection SDK class is present.
        internal bool HasSheetProtection;

        internal SLWorksheet(List<Color> ThemeColors, List<Color> IndexedColors, double ThemeDefaultColumnWidth,
            long ThemeDefaultColumnWidthInEMU, int MaxDigitWidth, List<double> ColumnStepSize,
            double CalculatedDefaultRowHeight)
        {
            ForceCustomRowColumnDimensionsSplitting = false;

            ActiveCell = new SLCellPoint(1, 1);

            SheetViews = new List<SLSheetView>();

            IsDoubleColumnWidth = false;
            SheetFormatProperties = new SLSheetFormatProperties(ThemeDefaultColumnWidth, ThemeDefaultColumnWidthInEMU,
                MaxDigitWidth, ColumnStepSize, CalculatedDefaultRowHeight);

            RowProperties = new Dictionary<int, SLRowProperties>();
            ColumnProperties = new Dictionary<int, SLColumnProperties>();
            Cells = new Dictionary<SLCellPoint, SLCell>();

            HasSheetProtection = false;
            SheetProtection = new SLSheetProtection();

            HasAutoFilter = false;
            AutoFilter = new SLAutoFilter();

            MergeCells = new List<SLMergeCell>();

            ConditionalFormattings = new List<SLConditionalFormatting>();
            ConditionalFormattings2010 = new List<SLConditionalFormatting2010>();

            DataValidations = new List<SLDataValidation>();
            DataValidationDisablePrompts = false;
            DataValidationXWindow = null;
            DataValidationYWindow = null;

            Hyperlinks = new List<SLHyperlink>();

            PageSettings = new SLPageSettings(ThemeColors, IndexedColors);

            RowBreaks = new Dictionary<int, SLBreak>();
            ColumnBreaks = new Dictionary<int, SLBreak>();

            DrawingId = string.Empty;
            NextWorksheetDrawingId = 2;
            Pictures = new List<SLPicture>();
            Charts = new List<SLChart>();

            InitializeBackgroundPictureStuff();

            LegacyDrawingId = string.Empty;
            Authors = new List<string>();
            Comments = new Dictionary<SLCellPoint, SLComment>();

            Tables = new List<SLTable>();

            SparklineGroups = new List<SLSparklineGroup>();
        }

        internal bool ForceCustomRowColumnDimensionsSplitting { get; set; }

        /// <summary>
        ///     The default is A1. Not going to get the active cell from the worksheet.
        ///     This is purely for setting (not getting). If not A1, then we'll do stuff.
        /// </summary>
        internal SLCellPoint ActiveCell { get; set; }

        internal List<SLSheetView> SheetViews { get; set; }

        internal bool IsDoubleColumnWidth { get; set; }
        internal SLSheetFormatProperties SheetFormatProperties { get; set; }

        // For posterity, here's a note about styles:
        // It's column style, then row style, then cell style. In increasing priority.
        // So if a cell has the default style, then we check if the row it belongs to
        // has a style. If the row also has no style (haha), then we check if the column
        // the cell belongs to has a style. If all are default, then the cell is truly
        // without any fashion sense.
        // This seems to be Excel's way of cascading styles, so we follow.

        internal Dictionary<int, SLRowProperties> RowProperties { get; set; }
        internal Dictionary<int, SLColumnProperties> ColumnProperties { get; set; }
        internal Dictionary<SLCellPoint, SLCell> Cells { get; set; }
        internal SLSheetProtection SheetProtection { get; set; }
        internal SLAutoFilter AutoFilter { get; set; }

        internal List<SLMergeCell> MergeCells { get; set; }

        internal List<SLConditionalFormatting> ConditionalFormattings { get; set; }
        internal List<SLConditionalFormatting2010> ConditionalFormattings2010 { get; set; }

        internal List<SLDataValidation> DataValidations { get; set; }
        internal bool DataValidationDisablePrompts { get; set; }
        internal uint? DataValidationXWindow { get; set; }
        internal uint? DataValidationYWindow { get; set; }

        internal List<SLHyperlink> Hyperlinks { get; set; }

        internal SLPageSettings PageSettings { get; set; }

        internal Dictionary<int, SLBreak> RowBreaks { get; set; }
        internal Dictionary<int, SLBreak> ColumnBreaks { get; set; }

        // use the reference ID of the Drawing class directly
        internal string DrawingId { get; set; }

        internal uint NextWorksheetDrawingId { get; set; }

        internal List<SLPicture> Pictures { get; set; }

        internal List<SLChart> Charts { get; set; }

        internal bool ToAppendBackgroundPicture { get; set; }
        internal string BackgroundPictureId { get; set; }

        /// <summary>
        ///     if null, then don't have to do anything
        /// </summary>
        internal bool? BackgroundPictureDataIsInFile { get; set; }

        internal string BackgroundPictureFileName { get; set; }
        internal byte[] BackgroundPictureByteData { get; set; }
        internal ImagePartType BackgroundPictureImagePartType { get; set; }

        // for cell comments
        internal string LegacyDrawingId { get; set; }
        internal List<string> Authors { get; set; }
        internal Dictionary<SLCellPoint, SLComment> Comments { get; set; }

        internal List<SLTable> Tables { get; set; }

        internal List<SLSparklineGroup> SparklineGroups { get; set; }

        internal void InitializeBackgroundPictureStuff()
        {
            BackgroundPictureId = string.Empty;
            BackgroundPictureDataIsInFile = null;
            BackgroundPictureFileName = string.Empty;
            BackgroundPictureByteData = new byte[1];
            BackgroundPictureImagePartType = ImagePartType.Bmp;
        }

        internal void ToggleCustomRowColumnDimension(bool IsCustom)
        {
            SheetFormatProperties.HasDefaultColumnWidth = IsCustom;
            if (IsCustom)
                SheetFormatProperties.CustomHeight = IsCustom;
            else
                SheetFormatProperties.CustomHeight = null;
        }

        internal void RefreshSparklineGroups()
        {
            for (var i = SparklineGroups.Count - 1; i >= 0; --i)
                if (SparklineGroups[i].Sparklines.Count == 0)
                    SparklineGroups.RemoveAt(i);
        }
    }
}