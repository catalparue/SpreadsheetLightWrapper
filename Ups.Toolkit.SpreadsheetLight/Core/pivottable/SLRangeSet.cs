using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.pivottable
{
    internal class SLRangeSet
    {
        internal SLRangeSet()
        {
            SetAllNull();
        }

        internal uint? FieldItemIndexPage1 { get; set; }
        internal uint? FieldItemIndexPage2 { get; set; }
        internal uint? FieldItemIndexPage3 { get; set; }
        internal uint? FieldItemIndexPage4 { get; set; }
        internal string Reference { get; set; }
        internal string Name { get; set; }
        internal string Sheet { get; set; }
        internal string Id { get; set; }

        private void SetAllNull()
        {
            FieldItemIndexPage1 = null;
            FieldItemIndexPage2 = null;
            FieldItemIndexPage3 = null;
            FieldItemIndexPage4 = null;
            Reference = "";
            Name = "";
            Sheet = "";
            Id = "";
        }

        internal void FromRangeSet(RangeSet rs)
        {
            SetAllNull();

            if (rs.FieldItemIndexPage1 != null) FieldItemIndexPage1 = rs.FieldItemIndexPage1.Value;
            if (rs.FieldItemIndexPage2 != null) FieldItemIndexPage2 = rs.FieldItemIndexPage2.Value;
            if (rs.FieldItemIndexPage3 != null) FieldItemIndexPage3 = rs.FieldItemIndexPage3.Value;
            if (rs.FieldItemIndexPage4 != null) FieldItemIndexPage4 = rs.FieldItemIndexPage4.Value;
            if (rs.Reference != null) Reference = rs.Reference.Value;
            if (rs.Name != null) Name = rs.Name.Value;
            if (rs.Sheet != null) Sheet = rs.Sheet.Value;
            if (rs.Id != null) Id = rs.Id.Value;
        }

        internal RangeSet ToRangeSet()
        {
            var rs = new RangeSet();
            if (FieldItemIndexPage1 != null) rs.FieldItemIndexPage1 = FieldItemIndexPage1.Value;
            if (FieldItemIndexPage2 != null) rs.FieldItemIndexPage2 = FieldItemIndexPage2.Value;
            if (FieldItemIndexPage3 != null) rs.FieldItemIndexPage3 = FieldItemIndexPage3.Value;
            if (FieldItemIndexPage4 != null) rs.FieldItemIndexPage4 = FieldItemIndexPage4.Value;
            if ((Reference != null) && (Reference.Length > 0)) rs.Reference = Reference;
            if ((Name != null) && (Name.Length > 0)) rs.Name = Name;
            if ((Sheet != null) && (Sheet.Length > 0)) rs.Sheet = Sheet;
            if ((Id != null) && (Id.Length > 0)) rs.Id = Id;

            return rs;
        }

        internal SLRangeSet Clone()
        {
            var rs = new SLRangeSet();
            rs.FieldItemIndexPage1 = FieldItemIndexPage1;
            rs.FieldItemIndexPage2 = FieldItemIndexPage2;
            rs.FieldItemIndexPage3 = FieldItemIndexPage3;
            rs.FieldItemIndexPage4 = FieldItemIndexPage4;
            rs.Reference = Reference;
            rs.Name = Name;
            rs.Sheet = Sheet;
            rs.Id = Id;

            return rs;
        }
    }
}