using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.table
{
    internal class SLIconFilter
    {
        internal SLIconFilter()
        {
            SetAllNull();
        }

        internal IconSetValues IconSet { get; set; }
        internal uint? IconId { get; set; }

        private void SetAllNull()
        {
            IconSet = IconSetValues.ThreeArrows;
            IconId = null;
        }

        internal void FromIconFilter(IconFilter icf)
        {
            SetAllNull();

            IconSet = icf.IconSet.Value;
            if (icf.IconId != null) IconId = icf.IconId.Value;
        }

        internal IconFilter ToIconFilter()
        {
            var icf = new IconFilter();
            icf.IconSet = IconSet;
            if (IconId != null) icf.IconId = IconId.Value;

            return icf;
        }

        internal SLIconFilter Clone()
        {
            var icf = new SLIconFilter();
            icf.IconSet = IconSet;
            icf.IconId = IconId;

            return icf;
        }
    }
}