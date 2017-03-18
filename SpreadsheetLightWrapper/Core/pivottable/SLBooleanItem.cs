using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.pivottable
{
    internal class SLBooleanItem
    {
        internal SLBooleanItem()
        {
            SetAllNull();
        }

        internal List<int> MemberPropertyIndexes { get; set; }

        internal bool Val { get; set; }
        internal bool? Unused { get; set; }
        internal bool? Calculated { get; set; }
        internal string Caption { get; set; }
        internal uint? PropertyCount { get; set; }

        private void SetAllNull()
        {
            MemberPropertyIndexes = new List<int>();

            Val = true;
            Unused = null;
            Calculated = null;
            Caption = "";
            PropertyCount = null;
        }

        internal void FromBooleanItem(BooleanItem bi)
        {
            SetAllNull();

            if (bi.Val != null) Val = bi.Val.Value;
            if (bi.Unused != null) Unused = bi.Unused.Value;
            if (bi.Calculated != null) Calculated = bi.Calculated.Value;
            if (bi.Caption != null) Caption = bi.Caption.Value;
            if (bi.PropertyCount != null) PropertyCount = bi.PropertyCount.Value;

            MemberPropertyIndex mpi;
            using (var oxr = OpenXmlReader.Create(bi))
            {
                while (oxr.Read())
                    if (oxr.ElementType == typeof(MemberPropertyIndex))
                    {
                        // 0 is the default value.
                        mpi = (MemberPropertyIndex) oxr.LoadCurrentElement();
                        if (mpi.Val != null) MemberPropertyIndexes.Add(mpi.Val.Value);
                        else MemberPropertyIndexes.Add(0);
                    }
            }
        }

        internal BooleanItem ToBooleanItem()
        {
            var bi = new BooleanItem();
            bi.Val = Val;
            if (Unused != null) bi.Unused = Unused.Value;
            if (Calculated != null) bi.Calculated = Calculated.Value;
            if ((Caption != null) && (Caption.Length > 0)) bi.Caption = Caption;
            if (PropertyCount != null) bi.PropertyCount = PropertyCount.Value;

            foreach (var i in MemberPropertyIndexes)
                if (i != 0) bi.Append(new MemberPropertyIndex {Val = i});
                else bi.Append(new MemberPropertyIndex());

            return bi;
        }

        internal SLBooleanItem Clone()
        {
            var bi = new SLBooleanItem();
            bi.Val = Val;
            bi.Unused = Unused;
            bi.Calculated = Calculated;
            bi.Caption = Caption;
            bi.PropertyCount = PropertyCount;

            bi.MemberPropertyIndexes = new List<int>();
            foreach (var i in MemberPropertyIndexes)
                bi.MemberPropertyIndexes.Add(i);

            return bi;
        }
    }
}