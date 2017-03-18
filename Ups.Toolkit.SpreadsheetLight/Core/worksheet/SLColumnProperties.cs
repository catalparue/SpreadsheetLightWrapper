using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Ups.Toolkit.SpreadsheetLight.Core.worksheet
{
    /// <summary>
    ///     Encapsulates properties and methods for columns. This simulates the DocumentFormat.OpenXml.Spreadsheet.Column
    ///     class.
    /// </summary>
    internal class SLColumnProperties
    {
        private double fWidth;

        // this doubles as customWidth
        internal bool HasWidth;

        internal double ThemeDefaultColumnWidth;
        internal long ThemeDefaultColumnWidthInEMU;

        /// <summary>
        ///     Initializes an instance of SLColumnProperties.
        /// </summary>
        internal SLColumnProperties(double ThemeDefaultColumnWidth, long ThemeDefaultColumnWidthInEMU, int MaxDigitWidth,
            List<double> ColumnStepSize)
        {
            this.MaxDigitWidth = MaxDigitWidth;
            listColumnStepSize = new List<double>();
            for (var i = 0; i < ColumnStepSize.Count; ++i)
                listColumnStepSize.Add(ColumnStepSize[i]);

            this.ThemeDefaultColumnWidth = ThemeDefaultColumnWidth;
            this.ThemeDefaultColumnWidthInEMU = ThemeDefaultColumnWidthInEMU;
            Width = ThemeDefaultColumnWidth;
            WidthInEMU = ThemeDefaultColumnWidthInEMU;
            HasWidth = false;

            StyleIndex = 0;
            Hidden = false;
            BestFit = false;
            Phonetic = false;
            OutlineLevel = 0;
            Collapsed = false;
        }

        internal bool IsEmpty
        {
            get
            {
                return !HasWidth && (StyleIndex == 0) && !Hidden
                       && !BestFit && !Phonetic && (OutlineLevel == 0)
                       && !Collapsed;
            }
        }

        internal int MaxDigitWidth { get; set; }
        internal List<double> listColumnStepSize { get; set; }
        // The column width. This is in number of characters of the width of the digit (0, 1, ... 9) with the maximum width, as rendered in the Normal style's font. The Normal style's font is typically the minor font in the default point size.
        internal double Width
        {
            get { return fWidth; }
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
                    WidthInEMU = iPixels*SLDocument.PixelToEMU;
                    fWidth = iWholeNumber + listColumnStepSize[iStep];
                    HasWidth = true;

                    BestFit = false;
                }
            }
        }

        internal long WidthInEMU { get; private set; }

        internal uint StyleIndex { get; set; }
        internal bool Hidden { get; set; }
        internal bool BestFit { get; set; }
        internal bool Phonetic { get; set; }
        internal byte OutlineLevel { get; set; }
        internal bool Collapsed { get; set; }

        internal string ToHash()
        {
            var sb = new StringBuilder();
            sb.AppendFormat("{0},", HasWidth);
            sb.AppendFormat("{0},", Width.ToString(CultureInfo.InvariantCulture));
            sb.AppendFormat("{0},", StyleIndex.ToString(CultureInfo.InvariantCulture));
            sb.AppendFormat("{0},", Hidden);
            sb.AppendFormat("{0},", BestFit);
            sb.AppendFormat("{0},", Phonetic);
            sb.AppendFormat("{0},", OutlineLevel.ToString(CultureInfo.InvariantCulture));
            sb.AppendFormat("{0}", Collapsed);

            return sb.ToString();
        }

        internal SLColumnProperties Clone()
        {
            var cp = new SLColumnProperties(ThemeDefaultColumnWidth, ThemeDefaultColumnWidthInEMU, MaxDigitWidth,
                listColumnStepSize);
            cp.HasWidth = HasWidth;
            cp.fWidth = fWidth;
            cp.WidthInEMU = WidthInEMU;
            cp.StyleIndex = StyleIndex;
            cp.Hidden = Hidden;
            cp.BestFit = BestFit;
            cp.Phonetic = Phonetic;
            cp.OutlineLevel = OutlineLevel;
            cp.Collapsed = Collapsed;

            return cp;
        }
    }
}