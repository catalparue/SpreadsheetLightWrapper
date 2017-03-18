using System;
using DocumentFormat.OpenXml.Spreadsheet;
using Ups.Toolkit.SpreadsheetLight.Core.misc;

namespace Ups.Toolkit.SpreadsheetLight.Core.worksheet
{
    internal class SLRowProperties
    {
        internal bool CustomHeight;
        private double fHeight;

        internal bool HasHeight;

        internal SLRowProperties(double DefaultRowHeight)
        {
            this.DefaultRowHeight = DefaultRowHeight;
            StyleIndex = 0;
            Height = DefaultRowHeight;
            CustomHeight = false;
            HasHeight = false;
            Hidden = false;
            OutlineLevel = 0;
            Collapsed = false;
            ThickTop = false;
            ThickBottom = false;
            ShowPhonetic = false;
        }

        internal bool IsEmpty
        {
            get
            {
                return (StyleIndex == 0) && !HasHeight && !Hidden
                       && (OutlineLevel == 0) && !Collapsed
                       && !ThickTop && !ThickBottom && !ShowPhonetic;
            }
        }

        internal double DefaultRowHeight { get; set; }

        internal uint StyleIndex { get; set; }
        // The row height in points.
        internal double Height
        {
            get { return fHeight; }
            set
            {
                var fModifiedRowHeight = value/SLDocument.RowHeightMultiple;
                fModifiedRowHeight = Math.Ceiling(fModifiedRowHeight)*SLDocument.RowHeightMultiple;

                HeightInEMU = (long) (fModifiedRowHeight*SLConstants.PointToEMU);

                fHeight = fModifiedRowHeight;
                CustomHeight = true;
                HasHeight = true;
            }
        }

        internal long HeightInEMU { get; private set; }

        internal bool Hidden { get; set; }
        internal byte OutlineLevel { get; set; }
        internal bool Collapsed { get; set; }
        internal bool ThickTop { get; set; }
        internal bool ThickBottom { get; set; }
        internal bool ShowPhonetic { get; set; }

        internal Row ToRow()
        {
            var r = new Row();
            if (StyleIndex > 0)
            {
                r.StyleIndex = StyleIndex;
                r.CustomFormat = true;
            }
            if (HasHeight)
                r.Height = Height;
            if (CustomHeight)
                r.CustomHeight = true;

            if (Hidden) r.Hidden = Hidden;
            if (OutlineLevel > 0) r.OutlineLevel = OutlineLevel;
            if (Collapsed) r.Collapsed = Collapsed;
            if (ThickTop) r.ThickTop = ThickTop;
            if (ThickBottom) r.ThickBot = ThickBottom;
            if (ShowPhonetic) r.ShowPhonetic = ShowPhonetic;

            return r;
        }

        internal SLRowProperties Clone()
        {
            var rp = new SLRowProperties(DefaultRowHeight);
            rp.DefaultRowHeight = DefaultRowHeight;
            rp.StyleIndex = StyleIndex;
            rp.HasHeight = HasHeight;
            rp.fHeight = fHeight;
            rp.HeightInEMU = HeightInEMU;
            rp.Hidden = Hidden;
            rp.OutlineLevel = OutlineLevel;
            rp.Collapsed = Collapsed;
            rp.ThickTop = ThickTop;
            rp.ThickBottom = ThickBottom;
            rp.ShowPhonetic = ShowPhonetic;

            return rp;
        }
    }
}