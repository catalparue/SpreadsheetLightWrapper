using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.pivottable
{
    internal class SLErrorItem
    {
        internal SLErrorItem()
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

        internal void FromErrorItem(ErrorItem ei)
        {
            SetAllNull();

            if (ei.Val != null) Val = ei.Val.Value;
            if (ei.Unused != null) Unused = ei.Unused.Value;
            if (ei.Calculated != null) Calculated = ei.Calculated.Value;
            if (ei.Caption != null) Caption = ei.Caption.Value;
            if (ei.PropertyCount != null) PropertyCount = ei.PropertyCount.Value;
            if (ei.FormatIndex != null) FormatIndex = ei.FormatIndex.Value;
            if (ei.BackgroundColor != null) BackgroundColor = ei.BackgroundColor.Value;
            if (ei.ForegroundColor != null) ForegroundColor = ei.ForegroundColor.Value;
            if (ei.Italic != null) Italic = ei.Italic.Value;
            if (ei.Underline != null) Underline = ei.Underline.Value;
            if (ei.Strikethrough != null) Strikethrough = ei.Strikethrough.Value;
            if (ei.Bold != null) Bold = ei.Bold.Value;

            SLTuplesType tt;
            MemberPropertyIndex mpi;
            using (var oxr = OpenXmlReader.Create(ei))
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

        internal ErrorItem ToErrorItem()
        {
            var ei = new ErrorItem();
            ei.Val = Val;
            if (Unused != null) ei.Unused = Unused.Value;
            if (Calculated != null) ei.Calculated = Calculated.Value;
            if ((Caption != null) && (Caption.Length > 0)) ei.Caption = Caption;
            if (PropertyCount != null) ei.PropertyCount = PropertyCount.Value;
            if (FormatIndex != null) ei.FormatIndex = FormatIndex.Value;
            if ((BackgroundColor != null) && (BackgroundColor.Length > 0))
                ei.BackgroundColor = new HexBinaryValue(BackgroundColor);
            if ((ForegroundColor != null) && (ForegroundColor.Length > 0))
                ei.ForegroundColor = new HexBinaryValue(ForegroundColor);
            if (Italic) ei.Italic = Italic;
            if (Underline) ei.Underline = Underline;
            if (Strikethrough) ei.Strikethrough = Strikethrough;
            if (Bold) ei.Bold = Bold;

            foreach (var tt in Tuples)
                ei.Append(tt.ToTuples());

            foreach (var i in MemberPropertyIndexes)
                if (i != 0) ei.Append(new MemberPropertyIndex {Val = i});
                else ei.Append(new MemberPropertyIndex());

            return ei;
        }

        internal SLErrorItem Clone()
        {
            var ei = new SLErrorItem();
            ei.Val = Val;
            ei.Unused = Unused;
            ei.Calculated = Calculated;
            ei.Caption = Caption;
            ei.PropertyCount = PropertyCount;
            ei.FormatIndex = FormatIndex;
            ei.BackgroundColor = BackgroundColor;
            ei.ForegroundColor = ForegroundColor;
            ei.Italic = Italic;
            ei.Underline = Underline;
            ei.Strikethrough = Strikethrough;
            ei.Bold = Bold;

            ei.Tuples = new List<SLTuplesType>();
            foreach (var tt in Tuples)
                ei.Tuples.Add(tt.Clone());

            ei.MemberPropertyIndexes = new List<int>();
            foreach (var i in MemberPropertyIndexes)
                ei.MemberPropertyIndexes.Add(i);

            return ei;
        }
    }
}