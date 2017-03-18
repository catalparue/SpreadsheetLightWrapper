using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.pivottable
{
    internal class SLStringItem
    {
        internal SLStringItem()
        {
            SetAllNull();
        }

        internal List<SLTuplesType> Tuples { get; set; }
        internal List<int> MemberPropertyIndexes { get; set; }

        internal string Val { get; set; }
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

            Val = "";
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

        internal void FromStringItem(StringItem si)
        {
            SetAllNull();

            if (si.Val != null) Val = si.Val.Value;
            if (si.Unused != null) Unused = si.Unused.Value;
            if (si.Calculated != null) Calculated = si.Calculated.Value;
            if (si.Caption != null) Caption = si.Caption.Value;
            if (si.PropertyCount != null) PropertyCount = si.PropertyCount.Value;
            if (si.FormatIndex != null) FormatIndex = si.FormatIndex.Value;
            if (si.BackgroundColor != null) BackgroundColor = si.BackgroundColor.Value;
            if (si.ForegroundColor != null) ForegroundColor = si.ForegroundColor.Value;
            if (si.Italic != null) Italic = si.Italic.Value;
            if (si.Underline != null) Underline = si.Underline.Value;
            if (si.Strikethrough != null) Strikethrough = si.Strikethrough.Value;
            if (si.Bold != null) Bold = si.Bold.Value;

            SLTuplesType tt;
            MemberPropertyIndex mpi;
            using (var oxr = OpenXmlReader.Create(si))
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

        internal StringItem ToStringItem()
        {
            var si = new StringItem();
            si.Val = Val;
            if (Unused != null) si.Unused = Unused.Value;
            if (Calculated != null) si.Calculated = Calculated.Value;
            if ((Caption != null) && (Caption.Length > 0)) si.Caption = Caption;
            if (PropertyCount != null) si.PropertyCount = PropertyCount.Value;
            if (FormatIndex != null) si.FormatIndex = FormatIndex.Value;
            if ((BackgroundColor != null) && (BackgroundColor.Length > 0))
                si.BackgroundColor = new HexBinaryValue(BackgroundColor);
            if ((ForegroundColor != null) && (ForegroundColor.Length > 0))
                si.ForegroundColor = new HexBinaryValue(ForegroundColor);
            if (Italic) si.Italic = Italic;
            if (Underline) si.Underline = Underline;
            if (Strikethrough) si.Strikethrough = Strikethrough;
            if (Bold) si.Bold = Bold;

            foreach (var tt in Tuples)
                si.Append(tt.ToTuples());

            foreach (var i in MemberPropertyIndexes)
                if (i != 0) si.Append(new MemberPropertyIndex {Val = i});
                else si.Append(new MemberPropertyIndex());

            return si;
        }

        internal SLStringItem Clone()
        {
            var si = new SLStringItem();
            si.Val = Val;
            si.Unused = Unused;
            si.Calculated = Calculated;
            si.Caption = Caption;
            si.PropertyCount = PropertyCount;
            si.FormatIndex = FormatIndex;
            si.BackgroundColor = BackgroundColor;
            si.ForegroundColor = ForegroundColor;
            si.Italic = Italic;
            si.Underline = Underline;
            si.Strikethrough = Strikethrough;
            si.Bold = Bold;

            si.Tuples = new List<SLTuplesType>();
            foreach (var tt in Tuples)
                si.Tuples.Add(tt.Clone());

            si.MemberPropertyIndexes = new List<int>();
            foreach (var i in MemberPropertyIndexes)
                si.MemberPropertyIndexes.Add(i);

            return si;
        }
    }
}