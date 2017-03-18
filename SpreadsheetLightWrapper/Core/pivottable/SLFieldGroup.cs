using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.pivottable
{
    internal class SLFieldGroup
    {
        internal bool HasGroupItems;
        internal bool HasRangeProperties;

        internal SLFieldGroup()
        {
            SetAllNull();
        }

        internal SLRangeProperties RangeProperties { get; set; }

        internal List<uint> DiscreteProperties { get; set; }
        internal SLGroupItems GroupItems { get; set; }

        internal uint? ParentId { get; set; }
        internal uint? Base { get; set; }

        private void SetAllNull()
        {
            HasRangeProperties = false;
            RangeProperties = new SLRangeProperties();

            DiscreteProperties = new List<uint>();

            HasGroupItems = false;
            GroupItems = new SLGroupItems();

            ParentId = null;
            Base = null;
        }

        internal void FromFieldGroup(FieldGroup fg)
        {
            SetAllNull();

            if (fg.ParentId != null) ParentId = fg.ParentId.Value;
            if (fg.Base != null) Base = fg.Base.Value;

            using (var oxr = OpenXmlReader.Create(fg))
            {
                while (oxr.Read())
                    if (oxr.ElementType == typeof(RangeProperties))
                    {
                        RangeProperties.FromRangeProperties((RangeProperties) oxr.LoadCurrentElement());
                        HasRangeProperties = true;
                    }
                    else if (oxr.ElementType == typeof(DiscreteProperties))
                    {
                        var dp = (DiscreteProperties) oxr.LoadCurrentElement();
                        FieldItem fi;
                        using (var oxrDiscrete = OpenXmlReader.Create(dp))
                        {
                            while (oxrDiscrete.Read())
                                if (oxrDiscrete.ElementType == typeof(FieldItem))
                                {
                                    fi = (FieldItem) oxrDiscrete.LoadCurrentElement();
                                    DiscreteProperties.Add(fi.Val);
                                }
                        }
                    }
                    else if (oxr.ElementType == typeof(GroupItems))
                    {
                        GroupItems.FromGroupItems((GroupItems) oxr.LoadCurrentElement());
                        HasGroupItems = true;
                    }
            }
        }

        internal FieldGroup ToFieldGroup()
        {
            var fg = new FieldGroup();
            if (ParentId != null) fg.ParentId = ParentId.Value;
            if (Base != null) fg.Base = Base.Value;

            if (HasRangeProperties)
                fg.Append(RangeProperties.ToRangeProperties());

            if (DiscreteProperties.Count > 0)
            {
                var dp = new DiscreteProperties();
                dp.Count = (uint) DiscreteProperties.Count;
                foreach (var i in DiscreteProperties)
                    dp.Append(new FieldItem {Val = i});

                fg.Append(dp);
            }

            if (HasGroupItems)
                fg.Append(GroupItems.ToGroupItems());

            return fg;
        }

        internal SLFieldGroup Clone()
        {
            var fg = new SLFieldGroup();
            fg.ParentId = ParentId;
            fg.Base = Base;

            fg.HasRangeProperties = HasRangeProperties;
            fg.RangeProperties = RangeProperties.Clone();

            fg.DiscreteProperties = new List<uint>();
            foreach (var i in DiscreteProperties)
                fg.DiscreteProperties.Add(i);

            fg.HasGroupItems = HasGroupItems;
            fg.GroupItems = GroupItems.Clone();

            return fg;
        }
    }
}