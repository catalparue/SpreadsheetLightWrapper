using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.pivottable
{
    internal class SLNumberItem
    {
        internal SLNumberItem()
        {
            SetAllNull();
        }

        internal List<SLTuplesType> Tuples { get; set; }
        internal List<int> MemberPropertyIndexes { get; set; }

        internal double Val { get; set; }
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

            Val = 0;
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

        internal void FromNumberItem(NumberItem ni)
        {
            SetAllNull();

            if (ni.Val != null) Val = ni.Val.Value;
            if (ni.Unused != null) Unused = ni.Unused.Value;
            if (ni.Calculated != null) Calculated = ni.Calculated.Value;
            if (ni.Caption != null) Caption = ni.Caption.Value;
            if (ni.PropertyCount != null) PropertyCount = ni.PropertyCount.Value;
            if (ni.FormatIndex != null) FormatIndex = ni.FormatIndex.Value;
            if (ni.BackgroundColor != null) BackgroundColor = ni.BackgroundColor.Value;
            if (ni.ForegroundColor != null) ForegroundColor = ni.ForegroundColor.Value;
            if (ni.Italic != null) Italic = ni.Italic.Value;
            if (ni.Underline != null) Underline = ni.Underline.Value;
            if (ni.Strikethrough != null) Strikethrough = ni.Strikethrough.Value;
            if (ni.Bold != null) Bold = ni.Bold.Value;

            SLTuplesType tt;
            MemberPropertyIndex mpi;
            using (var oxr = OpenXmlReader.Create(ni))
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

        internal NumberItem ToNumberItem()
        {
            var ni = new NumberItem();
            ni.Val = Val;
            if (Unused != null) ni.Unused = Unused.Value;
            if (Calculated != null) ni.Calculated = Calculated.Value;
            if ((Caption != null) && (Caption.Length > 0)) ni.Caption = Caption;
            if (PropertyCount != null) ni.PropertyCount = PropertyCount.Value;
            if (FormatIndex != null) ni.FormatIndex = FormatIndex.Value;
            if ((BackgroundColor != null) && (BackgroundColor.Length > 0))
                ni.BackgroundColor = new HexBinaryValue(BackgroundColor);
            if ((ForegroundColor != null) && (ForegroundColor.Length > 0))
                ni.ForegroundColor = new HexBinaryValue(ForegroundColor);
            if (Italic) ni.Italic = Italic;
            if (Underline) ni.Underline = Underline;
            if (Strikethrough) ni.Strikethrough = Strikethrough;
            if (Bold) ni.Bold = Bold;

            foreach (var tt in Tuples)
                ni.Append(tt.ToTuples());

            foreach (var i in MemberPropertyIndexes)
                if (i != 0) ni.Append(new MemberPropertyIndex {Val = i});
                else ni.Append(new MemberPropertyIndex());

            return ni;
        }

        internal SLNumberItem Clone()
        {
            var ni = new SLNumberItem();
            ni.Val = Val;
            ni.Unused = Unused;
            ni.Calculated = Calculated;
            ni.Caption = Caption;
            ni.PropertyCount = PropertyCount;
            ni.FormatIndex = FormatIndex;
            ni.BackgroundColor = BackgroundColor;
            ni.ForegroundColor = ForegroundColor;
            ni.Italic = Italic;
            ni.Underline = Underline;
            ni.Strikethrough = Strikethrough;
            ni.Bold = Bold;

            ni.Tuples = new List<SLTuplesType>();
            foreach (var tt in Tuples)
                ni.Tuples.Add(tt.Clone());

            ni.MemberPropertyIndexes = new List<int>();
            foreach (var i in MemberPropertyIndexes)
                ni.MemberPropertyIndexes.Add(i);

            return ni;
        }
    }
}