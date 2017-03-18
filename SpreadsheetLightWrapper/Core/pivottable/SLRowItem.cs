using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.pivottable
{
    internal class SLRowItem
    {
        internal SLRowItem()
        {
            SetAllNull();
        }

        internal List<int> MemberPropertyIndexes { get; set; }

        internal ItemValues ItemType { get; set; }
        internal uint RepeatedItemCount { get; set; }
        internal uint Index { get; set; }

        private void SetAllNull()
        {
            ItemType = ItemValues.Data;
            RepeatedItemCount = 0;
            Index = 0;
        }

        internal void FromRowItem(RowItem ri)
        {
            SetAllNull();

            if (ri.ItemType != null) ItemType = ri.ItemType.Value;
            if (ri.RepeatedItemCount != null) RepeatedItemCount = ri.RepeatedItemCount.Value;
            if (ri.Index != null) Index = ri.Index.Value;

            MemberPropertyIndex mpi;
            using (var oxr = OpenXmlReader.Create(ri))
            {
                while (oxr.Read())
                    if (oxr.ElementType == typeof(MemberPropertyIndex))
                    {
                        mpi = (MemberPropertyIndex) oxr.LoadCurrentElement();
                        if (mpi.Val != null) MemberPropertyIndexes.Add(mpi.Val.Value);
                        else MemberPropertyIndexes.Add(0);
                    }
            }
        }

        internal RowItem ToRowItem()
        {
            var ri = new RowItem();
            if (ItemType != ItemValues.Data) ri.ItemType = ItemType;
            if (RepeatedItemCount != 0) ri.RepeatedItemCount = RepeatedItemCount;
            if (Index != 0) ri.Index = Index;

            foreach (var i in MemberPropertyIndexes)
                if (i != 0) ri.Append(new MemberPropertyIndex {Val = i});
                else ri.Append(new MemberPropertyIndex());

            return ri;
        }

        internal SLRowItem Clone()
        {
            var ri = new SLRowItem();
            ri.ItemType = ItemType;
            ri.RepeatedItemCount = RepeatedItemCount;
            ri.Index = Index;

            ri.MemberPropertyIndexes = new List<int>();
            foreach (var i in MemberPropertyIndexes)
                ri.MemberPropertyIndexes.Add(i);

            return ri;
        }
    }
}