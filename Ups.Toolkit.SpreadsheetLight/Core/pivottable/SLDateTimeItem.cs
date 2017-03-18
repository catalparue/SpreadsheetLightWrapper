using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.pivottable
{
    internal class SLDateTimeItem
    {
        internal SLDateTimeItem()
        {
            SetAllNull();
        }

        internal List<int> MemberPropertyIndexes { get; set; }

        internal DateTime Val { get; set; }
        internal bool? Unused { get; set; }
        internal bool? Calculated { get; set; }
        internal string Caption { get; set; }
        internal uint? PropertyCount { get; set; }

        private void SetAllNull()
        {
            MemberPropertyIndexes = new List<int>();

            Val = new DateTime();
            Unused = null;
            Calculated = null;
            Caption = "";
            PropertyCount = null;
        }

        internal void FromDateTimeItem(DateTimeItem dti)
        {
            SetAllNull();

            if (dti.Val != null) Val = dti.Val.Value;
            if (dti.Unused != null) Unused = dti.Unused.Value;
            if (dti.Calculated != null) Calculated = dti.Calculated.Value;
            if (dti.Caption != null) Caption = dti.Caption.Value;
            if (dti.PropertyCount != null) PropertyCount = dti.PropertyCount.Value;

            MemberPropertyIndex mpi;
            using (var oxr = OpenXmlReader.Create(dti))
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

        internal DateTimeItem ToDateTimeItem()
        {
            var dti = new DateTimeItem();
            dti.Val = Val;
            if (Unused != null) dti.Unused = Unused.Value;
            if (Calculated != null) dti.Calculated = Calculated.Value;
            if ((Caption != null) && (Caption.Length > 0)) dti.Caption = Caption;
            if (PropertyCount != null) dti.PropertyCount = PropertyCount.Value;

            foreach (var i in MemberPropertyIndexes)
                if (i != 0) dti.Append(new MemberPropertyIndex {Val = i});
                else dti.Append(new MemberPropertyIndex());

            return dti;
        }

        internal SLDateTimeItem Clone()
        {
            var dti = new SLDateTimeItem();
            dti.Val = Val;
            dti.Unused = Unused;
            dti.Calculated = Calculated;
            dti.Caption = Caption;
            dti.PropertyCount = PropertyCount;

            dti.MemberPropertyIndexes = new List<int>();
            foreach (var i in MemberPropertyIndexes)
                dti.MemberPropertyIndexes.Add(i);

            return dti;
        }
    }
}