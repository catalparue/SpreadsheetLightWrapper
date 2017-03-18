using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.pivottable
{
    internal class SLMissingItem
    {
        internal SLMissingItem()
        {
            SetAllNull();
        }

        internal List<SLTuplesType> Tuples { get; set; }
        internal List<int> MemberPropertyIndexes { get; set; }

        internal bool? Unused { get; set; }
        internal bool? Calculated { get; set; }
        internal string Caption { get; set; }
        internal uint? PropertyCount { get; set; }
        internal uint? FormatIndex { get; set; }
        internal string BackgroundColor { get; set; }
        internal string ForegroundColor { get; set; }
        internal bool Italic { get; set; }
        internal bool Underline { get; set; }
        internal bool Strikethrough { get; set; }
        internal bool Bold { get; set; }

        private void SetAllNull()
        {
            Tuples = new List<SLTuplesType>();
            MemberPropertyIndexes = new List<int>();

            Unused = null;
            Calculated = null;
            Caption = "";
            PropertyCount = null;
            FormatIndex = null;
            BackgroundColor = "";
            ForegroundColor = "";
            Italic = false;
            Underline = false;
            Strikethrough = false;
            Bold = false;
        }

        internal void FromMissingItem(MissingItem mi)
        {
            SetAllNull();

            if (mi.Unused != null) Unused = mi.Unused.Value;
            if (mi.Calculated != null) Calculated = mi.Calculated.Value;
            if (mi.Caption != null) Caption = mi.Caption.Value;
            if (mi.PropertyCount != null) PropertyCount = mi.PropertyCount.Value;
            if (mi.FormatIndex != null) FormatIndex = mi.FormatIndex.Value;
            if (mi.BackgroundColor != null) BackgroundColor = mi.BackgroundColor.Value;
            if (mi.ForegroundColor != null) ForegroundColor = mi.ForegroundColor.Value;
            if (mi.Italic != null) Italic = mi.Italic.Value;
            if (mi.Underline != null) Underline = mi.Underline.Value;
            if (mi.Strikethrough != null) Strikethrough = mi.Strikethrough.Value;
            if (mi.Bold != null) Bold = mi.Bold.Value;

            SLTuplesType tt;
            MemberPropertyIndex mpi;
            using (var oxr = OpenXmlReader.Create(mi))
            {
                while (oxr.Read())
                    if (oxr.ElementType == typeof(Tuples))
                    {
                        tt = new SLTuplesType();
                        tt.FromTuples((Tuples) oxr.LoadCurrentElement());
                        Tuples.Add(tt);
                    }
                    else if (oxr.ElementType == typeof(MemberPropertyIndex))
                    {
                        // 0 is the default value.
                        mpi = (MemberPropertyIndex) oxr.LoadCurrentElement();
                        if (mpi.Val != null) MemberPropertyIndexes.Add(mpi.Val.Value);
                        else MemberPropertyIndexes.Add(0);
                    }
            }
        }

        internal MissingItem ToMissingItem()
        {
            var mi = new MissingItem();
            if (Unused != null) mi.Unused = Unused.Value;
            if (Calculated != null) mi.Calculated = Calculated.Value;
            if ((Caption != null) && (Caption.Length > 0)) mi.Caption = Caption;
            if (PropertyCount != null) mi.PropertyCount = PropertyCount.Value;
            if (FormatIndex != null) mi.FormatIndex = FormatIndex.Value;
            if ((BackgroundColor != null) && (BackgroundColor.Length > 0))
                mi.BackgroundColor = new HexBinaryValue(BackgroundColor);
            if ((ForegroundColor != null) && (ForegroundColor.Length > 0))
                mi.ForegroundColor = new HexBinaryValue(ForegroundColor);
            if (Italic) mi.Italic = Italic;
            if (Underline) mi.Underline = Underline;
            if (Strikethrough) mi.Strikethrough = Strikethrough;
            if (Bold) mi.Bold = Bold;

            foreach (var tt in Tuples)
                mi.Append(tt.ToTuples());

            foreach (var i in MemberPropertyIndexes)
                if (i != 0) mi.Append(new MemberPropertyIndex {Val = i});
                else mi.Append(new MemberPropertyIndex());

            return mi;
        }

        internal SLMissingItem Clone()
        {
            var mi = new SLMissingItem();
            mi.Unused = Unused;
            mi.Calculated = Calculated;
            mi.Caption = Caption;
            mi.PropertyCount = PropertyCount;
            mi.FormatIndex = FormatIndex;
            mi.BackgroundColor = BackgroundColor;
            mi.ForegroundColor = ForegroundColor;
            mi.Italic = Italic;
            mi.Underline = Underline;
            mi.Strikethrough = Strikethrough;
            mi.Bold = Bold;

            mi.Tuples = new List<SLTuplesType>();
            foreach (var tt in Tuples)
                mi.Tuples.Add(tt.Clone());

            mi.MemberPropertyIndexes = new List<int>();
            foreach (var i in MemberPropertyIndexes)
                mi.MemberPropertyIndexes.Add(i);

            return mi;
        }
    }
}