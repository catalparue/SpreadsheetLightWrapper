using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace Ups.Toolkit.SpreadsheetLight.Core.office2010
{
    internal class SLConditionalFormattingIcon2010
    {
        internal SLConditionalFormattingIcon2010()
        {
            SetAllNull();
        }

        //http://msdn.microsoft.com/en-us/library/documentformat.openxml.office2010.excel.conditionalformattingicon.aspx

        internal X14.IconSetTypeValues IconSet { get; set; }
        internal uint IconId { get; set; }

        private void SetAllNull()
        {
            IconSet = X14.IconSetTypeValues.ThreeTrafficLights1;
            IconId = 0;
        }

        internal void FromConditionalFormattingIcon(X14.ConditionalFormattingIcon cfi)
        {
            SetAllNull();

            if (cfi.IconSet != null) IconSet = cfi.IconSet.Value;
            if (cfi.IconId != null) IconId = cfi.IconId.Value;
        }

        internal X14.ConditionalFormattingIcon ToConditionalFormattingIcon()
        {
            var cfi = new X14.ConditionalFormattingIcon();

            cfi.IconSet = IconSet;
            cfi.IconId = IconId;

            return cfi;
        }

        internal SLConditionalFormattingIcon2010 Clone()
        {
            var cfi = new SLConditionalFormattingIcon2010();
            cfi.IconSet = IconSet;
            cfi.IconId = IconId;

            return cfi;
        }
    }
}