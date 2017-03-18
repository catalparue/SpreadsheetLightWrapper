using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;
using Ups.Toolkit.SpreadsheetLight.Core.misc;

namespace Ups.Toolkit.SpreadsheetLight.Core.worksheet
{
    internal class SLSheetFormatProperties
    {
        internal double CalculatedDefaultRowHeight;

        internal double fDefaultColumnWidth;

        internal double fDefaultRowHeight;

        internal long lDefaultColumnWidthInEMU;

        internal long lDefaultRowHeightInEMU;
        internal double ThemeDefaultColumnWidth;
        internal long ThemeDefaultColumnWidthInEMU;

        internal SLSheetFormatProperties(double ThemeDefaultColumnWidth, long ThemeDefaultColumnWidthInEMU,
            int MaxDigitWidth, List<double> ColumnStepSize, double CalculatedDefaultRowHeight)
        {
            this.MaxDigitWidth = MaxDigitWidth;
            listColumnStepSize = new List<double>();
            for (var i = 0; i < ColumnStepSize.Count; ++i)
                listColumnStepSize.Add(ColumnStepSize[i]);

            BaseColumnWidth = null;

            this.ThemeDefaultColumnWidth = ThemeDefaultColumnWidth;
            this.ThemeDefaultColumnWidthInEMU = ThemeDefaultColumnWidthInEMU;
            fDefaultColumnWidth = ThemeDefaultColumnWidth;
            lDefaultColumnWidthInEMU = ThemeDefaultColumnWidthInEMU;
            HasDefaultColumnWidth = false;

            this.CalculatedDefaultRowHeight = CalculatedDefaultRowHeight;
            fDefaultRowHeight = CalculatedDefaultRowHeight;
            lDefaultRowHeightInEMU = Convert.ToInt64(CalculatedDefaultRowHeight*SLConstants.PointToEMU);

            CustomHeight = null;
            ZeroHeight = null;
            ThickTop = null;
            ThickBottom = null;
            OutlineLevelRow = null;
            OutlineLevelColumn = null;
        }

        internal int MaxDigitWidth { get; set; }
        internal List<double> listColumnStepSize { get; set; }

        internal uint? BaseColumnWidth { get; set; }

        internal bool HasDefaultColumnWidth { get; set; }

        internal double DefaultColumnWidth
        {
            get { return fDefaultColumnWidth; }
            set
            {
                var fValue = value;
                if (fValue > 0)
                {
                    var iWholeNumber = Convert.ToInt32(Math.Truncate(fValue));
                    var fRemainder = fValue - iWholeNumber;

                    var iStep = 0;
                    for (iStep = listColumnStepSize.Count - 1; iStep >= 0; --iStep)
                        if (fRemainder > listColumnStepSize[iStep]) break;

                    // this is in case (fRemainder > listColumnStepSize[iStep]) evaluates
                    // to false when fRemainder is 0.0 and listColumnStepSize[0] is also 0.0
                    // and I hate checking for equality between floating point values...
                    // By then iStep should be -1, which breaks the loop.
                    if (iStep < 0) iStep = 0;

                    // the step sizes were calculated based on the max digit width minus 1 pixel.
                    var iPixels = iWholeNumber*(MaxDigitWidth - 1) + iStep;
                    lDefaultColumnWidthInEMU = iPixels*SLDocument.PixelToEMU;
                    fDefaultColumnWidth = iWholeNumber + listColumnStepSize[iStep];
                    HasDefaultColumnWidth = true;
                }
            }
        }

        internal long DefaultColumnWidthInEMU
        {
            get { return lDefaultColumnWidthInEMU; }
        }

        internal double DefaultRowHeight
        {
            get { return fDefaultRowHeight; }
            set
            {
                var fModifiedRowHeight = value/SLDocument.RowHeightMultiple;
                // round because it looks nicer. Is 4 decimal places good enough?
                fModifiedRowHeight = Math.Round(Math.Ceiling(fModifiedRowHeight)*SLDocument.RowHeightMultiple, 4);

                lDefaultRowHeightInEMU = (long) (fModifiedRowHeight*SLConstants.PointToEMU);

                fDefaultRowHeight = fModifiedRowHeight;
            }
        }

        internal long DefaultRowHeightInEMU
        {
            get { return lDefaultRowHeightInEMU; }
        }

        internal bool? CustomHeight { get; set; }
        internal bool? ZeroHeight { get; set; }
        internal bool? ThickTop { get; set; }
        internal bool? ThickBottom { get; set; }
        internal byte? OutlineLevelRow { get; set; }
        internal byte? OutlineLevelColumn { get; set; }

        internal void FromSheetFormatProperties(SheetFormatProperties sfp)
        {
            if (sfp.BaseColumnWidth != null) BaseColumnWidth = sfp.BaseColumnWidth.Value;
            else BaseColumnWidth = null;

            if (sfp.DefaultColumnWidth != null)
            {
                DefaultColumnWidth = sfp.DefaultColumnWidth.Value;
                HasDefaultColumnWidth = true;
            }
            else
            {
                fDefaultColumnWidth = ThemeDefaultColumnWidth;
                lDefaultRowHeightInEMU = ThemeDefaultColumnWidthInEMU;
                HasDefaultColumnWidth = false;
            }

            if (sfp.DefaultRowHeight != null)
            {
                DefaultRowHeight = sfp.DefaultRowHeight.Value;
            }
            else
            {
                fDefaultRowHeight = CalculatedDefaultRowHeight;
                lDefaultRowHeightInEMU = Convert.ToInt64(CalculatedDefaultRowHeight*SLConstants.PointToEMU);
            }

            if (sfp.CustomHeight != null) CustomHeight = sfp.CustomHeight.Value;
            else CustomHeight = null;

            if (sfp.ZeroHeight != null) ZeroHeight = sfp.ZeroHeight.Value;
            else ZeroHeight = null;

            if (sfp.ThickTop != null) ThickTop = sfp.ThickTop.Value;
            else ThickTop = null;

            if (sfp.ThickBottom != null) ThickBottom = sfp.ThickBottom.Value;
            else ThickBottom = null;

            if (sfp.OutlineLevelRow != null) OutlineLevelRow = sfp.OutlineLevelRow.Value;
            else OutlineLevelRow = null;

            if (sfp.OutlineLevelColumn != null) OutlineLevelColumn = sfp.OutlineLevelColumn.Value;
            else OutlineLevelColumn = null;
        }

        internal SheetFormatProperties ToSheetFormatProperties()
        {
            var sfp = new SheetFormatProperties();
            if (BaseColumnWidth != null) sfp.BaseColumnWidth = BaseColumnWidth.Value;

            if (HasDefaultColumnWidth)
                sfp.DefaultColumnWidth = DefaultColumnWidth;

            sfp.DefaultRowHeight = DefaultRowHeight;

            if (CustomHeight != null) sfp.CustomHeight = CustomHeight.Value;
            if (ZeroHeight != null) sfp.ZeroHeight = ZeroHeight.Value;
            if (ThickTop != null) sfp.ThickTop = ThickTop.Value;
            if (ThickBottom != null) sfp.ThickBottom = ThickBottom.Value;
            if (OutlineLevelRow != null) sfp.OutlineLevelRow = OutlineLevelRow.Value;
            if (OutlineLevelColumn != null) sfp.OutlineLevelColumn = OutlineLevelColumn.Value;

            return sfp;
        }
    }
}