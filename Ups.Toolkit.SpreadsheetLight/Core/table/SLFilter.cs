using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.table
{
    internal class SLFilter
    {
        internal SLFilter()
        {
            Val = string.Empty;
        }

        internal string Val { get; set; }

        internal void FromFilter(Filter f)
        {
            Val = f.Val ?? string.Empty;
        }

        internal Filter ToFilter()
        {
            var f = new Filter();
            f.Val = Val;

            return f;
        }

        internal SLFilter Clone()
        {
            var f = new SLFilter();
            f.Val = Val;

            return f;
        }
    }
}