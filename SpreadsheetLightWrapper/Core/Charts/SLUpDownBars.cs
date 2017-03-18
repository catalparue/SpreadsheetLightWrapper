using System.Collections.Generic;
using System.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLightWrapper.Core.Charts
{
    /// <summary>
    ///     Encapsulates properties and methods for up-down bars.
    ///     This simulates the DocumentFormat.OpenXml.Drawing.Charts.UpDownBars class.
    /// </summary>
    public class SLUpDownBars
    {
        internal ushort iGapWidth;

        internal SLUpDownBars(List<Color> ThemeColors, bool IsStylish = false)
        {
            iGapWidth = 150;
            UpBars = new SLUpBars(ThemeColors, IsStylish);
            DownBars = new SLDownBars(ThemeColors, IsStylish);
        }

        /// <summary>
        ///     The gap width between consecutive up-down bars as a percentage of the width of the bar, ranging from 0 to 500 (both
        ///     inclusive).
        /// </summary>
        public ushort GapWidth
        {
            get { return iGapWidth; }
            set
            {
                iGapWidth = value;
                if (iGapWidth > 500) iGapWidth = 500;
            }
        }

        /// <summary>
        ///     The up bars.
        /// </summary>
        public SLUpBars UpBars { get; set; }

        /// <summary>
        ///     The down bars.
        /// </summary>
        public SLDownBars DownBars { get; set; }

        internal C.UpDownBars ToUpDownBars(bool IsStylish = false)
        {
            var udb = new C.UpDownBars();
            udb.GapWidth = new C.GapWidth {Val = iGapWidth};
            udb.UpBars = UpBars.ToUpBars(IsStylish);
            udb.DownBars = DownBars.ToDownBars(IsStylish);

            return udb;
        }

        internal SLUpDownBars Clone()
        {
            var udb = new SLUpDownBars(new List<Color>());
            udb.iGapWidth = iGapWidth;
            udb.UpBars = UpBars.Clone();
            udb.DownBars = DownBars.Clone();

            return udb;
        }
    }
}