using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.pivottable
{
    internal class SLItem
    {
        internal SLItem()
        {
            SetAllNull();
        }

        /// <summary>
        ///     Attribute: n
        /// </summary>
        internal string ItemName { get; set; }

        /// <summary>
        ///     Attribute: t
        /// </summary>
        internal ItemValues ItemType { get; set; }

        /// <summary>
        ///     Attribute: h
        /// </summary>
        internal bool Hidden { get; set; }

        /// <summary>
        ///     Attribute: s
        /// </summary>
        internal bool HasStringVlue { get; set; } // [sic]

        /// <summary>
        ///     Attribute: sd
        /// </summary>
        internal bool HideDetails { get; set; }

        /// <summary>
        ///     Attribute: f
        /// </summary>
        internal bool Calculated { get; set; }

        /// <summary>
        ///     Attribute: m
        /// </summary>
        internal bool Missing { get; set; }

        /// <summary>
        ///     Attribute: c
        /// </summary>
        internal bool ChildItems { get; set; }

        /// <summary>
        ///     Attribute: x
        /// </summary>
        internal uint? Index { get; set; }

        /// <summary>
        ///     Attribute: d
        /// </summary>
        internal bool Expanded { get; set; }

        /// <summary>
        ///     Attribute: e
        /// </summary>
        internal bool DrillAcrossAttributes { get; set; }

        private void SetAllNull()
        {
            ItemName = ""; //n
            ItemType = ItemValues.Data; //t
            Hidden = false; //h
            HasStringVlue = false; //s
            HideDetails = true; //sd
            Calculated = false; //f
            Missing = false; //m
            ChildItems = false; //c
            Index = null; //uint opt x
            Expanded = false; //d
            DrillAcrossAttributes = true; //e
        }

        internal void FromItem(Item it)
        {
            SetAllNull();

            if (it.ItemName != null) ItemName = it.ItemName.Value;
            if (it.ItemType != null) ItemType = it.ItemType.Value;
            if (it.Hidden != null) Hidden = it.Hidden.Value;
            if (it.HasStringVlue != null) HasStringVlue = it.HasStringVlue.Value;
            if (it.HideDetails != null) HideDetails = it.HideDetails.Value;
            if (it.Calculated != null) Calculated = it.Calculated.Value;
            if (it.Missing != null) Missing = it.Missing.Value;
            if (it.ChildItems != null) ChildItems = it.ChildItems.Value;
            if (it.Index != null) Index = it.Index.Value;
            if (it.Expanded != null) Expanded = it.Expanded.Value;
            if (it.DrillAcrossAttributes != null) DrillAcrossAttributes = it.DrillAcrossAttributes.Value;
        }

        internal Item ToItem()
        {
            var it = new Item();
            if (ItemName.Length > 0) it.ItemName = ItemName;
            if (ItemType != ItemValues.Data) it.ItemType = ItemType;
            if (Hidden) it.Hidden = Hidden;
            if (HasStringVlue) it.HasStringVlue = HasStringVlue;
            if (HideDetails != true) it.HideDetails = HideDetails;
            if (Calculated) it.Calculated = Calculated;
            if (Missing) it.Missing = Missing;
            if (ChildItems) it.ChildItems = ChildItems;
            if (Index != null) it.Index = Index.Value;
            if (Expanded) it.Expanded = Expanded;
            if (DrillAcrossAttributes != true) it.DrillAcrossAttributes = DrillAcrossAttributes;

            return it; // haha return it... maybe name a variable called "what"...
        }

        internal SLItem Clone()
        {
            var it = new SLItem();
            it.ItemName = ItemName;
            it.ItemType = ItemType;
            it.Hidden = Hidden;
            it.HasStringVlue = HasStringVlue;
            it.HideDetails = HideDetails;
            it.Calculated = Calculated;
            it.Missing = Missing;
            it.ChildItems = ChildItems;
            it.Index = Index;
            it.Expanded = Expanded;
            it.DrillAcrossAttributes = DrillAcrossAttributes;

            return it;
        }
    }
}