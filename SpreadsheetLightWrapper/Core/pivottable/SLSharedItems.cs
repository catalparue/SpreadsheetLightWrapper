using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.pivottable
{
    internal enum SLSharedGroupItemsType
    {
        Missing = 0,
        Number,
        Boolean,
        Error,
        String,
        DateTime
    }

    internal struct SLSharedGroupItemsTypeIndexPair
    {
        internal SLSharedGroupItemsType Type;
        // this is the 0-based index into whichever List<> depending on Type.
        internal int Index;

        internal SLSharedGroupItemsTypeIndexPair(SLSharedGroupItemsType Type, int Index)
        {
            this.Type = Type;
            this.Index = Index;
        }
    }

    internal class SLSharedItems
    {
        internal SLSharedItems()
        {
            SetAllNull();
        }

        internal List<SLSharedGroupItemsTypeIndexPair> Items { get; set; }

        internal List<SLMissingItem> MissingItems { get; set; }
        internal List<SLNumberItem> NumberItems { get; set; }
        internal List<SLBooleanItem> BooleanItems { get; set; }
        internal List<SLErrorItem> ErrorItems { get; set; }
        internal List<SLStringItem> StringItems { get; set; }
        internal List<SLDateTimeItem> DateTimeItems { get; set; }

        internal bool ContainsSemiMixedTypes { get; set; }
        internal bool ContainsNonDate { get; set; }
        internal bool ContainsDate { get; set; }
        internal bool ContainsString { get; set; }
        internal bool ContainsBlank { get; set; }
        internal bool ContainsMixedTypes { get; set; }
        internal bool ContainsNumber { get; set; }
        internal bool ContainsInteger { get; set; }
        internal double? MinValue { get; set; }
        internal double? MaxValue { get; set; }
        internal DateTime? MinDate { get; set; }
        internal DateTime? MaxDate { get; set; }
        //No need? internal uint? Count { get; set; }
        internal bool LongText { get; set; }

        private void SetAllNull()
        {
            Items = new List<SLSharedGroupItemsTypeIndexPair>();

            MissingItems = new List<SLMissingItem>();
            NumberItems = new List<SLNumberItem>();
            BooleanItems = new List<SLBooleanItem>();
            ErrorItems = new List<SLErrorItem>();
            StringItems = new List<SLStringItem>();
            DateTimeItems = new List<SLDateTimeItem>();

            ContainsSemiMixedTypes = true;
            ContainsNonDate = true;
            ContainsDate = false;
            ContainsString = true;
            ContainsBlank = false;
            ContainsMixedTypes = false;
            ContainsNumber = false;
            ContainsInteger = false;
            MinValue = null;
            MaxValue = null;
            MinDate = null;
            MaxDate = null;
            //this.Count = null;
            LongText = false;
        }

        internal void FromSharedItems(SharedItems sis)
        {
            SetAllNull();

            if (sis.ContainsSemiMixedTypes != null) ContainsSemiMixedTypes = sis.ContainsSemiMixedTypes.Value;
            if (sis.ContainsNonDate != null) ContainsNonDate = sis.ContainsNonDate.Value;
            if (sis.ContainsDate != null) ContainsDate = sis.ContainsDate.Value;
            if (sis.ContainsString != null) ContainsString = sis.ContainsString.Value;
            if (sis.ContainsBlank != null) ContainsBlank = sis.ContainsBlank.Value;
            if (sis.ContainsMixedTypes != null) ContainsMixedTypes = sis.ContainsMixedTypes.Value;
            if (sis.ContainsNumber != null) ContainsNumber = sis.ContainsNumber.Value;
            if (sis.ContainsInteger != null) ContainsInteger = sis.ContainsInteger.Value;
            if (sis.MinValue != null) MinValue = sis.MinValue.Value;
            if (sis.MaxValue != null) MaxValue = sis.MaxValue.Value;
            if (sis.MinDate != null) MinDate = sis.MinDate.Value;
            if (sis.MaxDate != null) MaxDate = sis.MaxDate.Value;
            //count
            if (sis.LongText != null) LongText = sis.LongText.Value;

            SLMissingItem mi;
            SLNumberItem ni;
            SLBooleanItem bi;
            SLErrorItem ei;
            SLStringItem si;
            SLDateTimeItem dti;
            using (var oxr = OpenXmlReader.Create(sis))
            {
                while (oxr.Read())
                    if (oxr.ElementType == typeof(MissingItem))
                    {
                        mi = new SLMissingItem();
                        mi.FromMissingItem((MissingItem) oxr.LoadCurrentElement());
                        Items.Add(new SLSharedGroupItemsTypeIndexPair(SLSharedGroupItemsType.Missing, MissingItems.Count));
                        MissingItems.Add(mi);
                    }
                    else if (oxr.ElementType == typeof(NumberItem))
                    {
                        ni = new SLNumberItem();
                        ni.FromNumberItem((NumberItem) oxr.LoadCurrentElement());
                        Items.Add(new SLSharedGroupItemsTypeIndexPair(SLSharedGroupItemsType.Number, NumberItems.Count));
                        NumberItems.Add(ni);
                    }
                    else if (oxr.ElementType == typeof(BooleanItem))
                    {
                        bi = new SLBooleanItem();
                        bi.FromBooleanItem((BooleanItem) oxr.LoadCurrentElement());
                        Items.Add(new SLSharedGroupItemsTypeIndexPair(SLSharedGroupItemsType.Boolean, BooleanItems.Count));
                        BooleanItems.Add(bi);
                    }
                    else if (oxr.ElementType == typeof(ErrorItem))
                    {
                        ei = new SLErrorItem();
                        ei.FromErrorItem((ErrorItem) oxr.LoadCurrentElement());
                        Items.Add(new SLSharedGroupItemsTypeIndexPair(SLSharedGroupItemsType.Error, ErrorItems.Count));
                        ErrorItems.Add(ei);
                    }
                    else if (oxr.ElementType == typeof(StringItem))
                    {
                        si = new SLStringItem();
                        si.FromStringItem((StringItem) oxr.LoadCurrentElement());
                        Items.Add(new SLSharedGroupItemsTypeIndexPair(SLSharedGroupItemsType.String, StringItems.Count));
                        StringItems.Add(si);
                    }
                    else if (oxr.ElementType == typeof(DateTimeItem))
                    {
                        dti = new SLDateTimeItem();
                        dti.FromDateTimeItem((DateTimeItem) oxr.LoadCurrentElement());
                        Items.Add(new SLSharedGroupItemsTypeIndexPair(SLSharedGroupItemsType.DateTime,
                            DateTimeItems.Count));
                        DateTimeItems.Add(dti);
                    }
            }
        }

        internal SharedItems ToSharedItems()
        {
            var sis = new SharedItems();
            if (ContainsSemiMixedTypes != true) sis.ContainsSemiMixedTypes = ContainsSemiMixedTypes;
            if (ContainsNonDate != true) sis.ContainsNonDate = ContainsNonDate;
            if (ContainsDate) sis.ContainsDate = ContainsDate;
            if (ContainsString != true) sis.ContainsString = ContainsString;
            if (ContainsBlank) sis.ContainsBlank = ContainsBlank;
            if (ContainsMixedTypes) sis.ContainsMixedTypes = ContainsMixedTypes;
            if (ContainsNumber) sis.ContainsNumber = ContainsNumber;
            if (ContainsInteger) sis.ContainsInteger = ContainsInteger;
            if (MinValue != null) sis.MinValue = MinValue.Value;
            if (MaxValue != null) sis.MaxValue = MaxValue.Value;
            if (MinDate != null) sis.MinDate = new DateTimeValue(MinDate.Value);
            if (MaxDate != null) sis.MaxDate = new DateTimeValue(MaxDate.Value);

            // is it the sum of all the various items?
            if (Items.Count > 0) sis.Count = (uint) Items.Count;

            if (LongText) sis.LongText = LongText;

            foreach (var pair in Items)
                switch (pair.Type)
                {
                    case SLSharedGroupItemsType.Missing:
                        sis.Append(MissingItems[pair.Index].ToMissingItem());
                        break;
                    case SLSharedGroupItemsType.Number:
                        sis.Append(NumberItems[pair.Index].ToNumberItem());
                        break;
                    case SLSharedGroupItemsType.Boolean:
                        sis.Append(BooleanItems[pair.Index].ToBooleanItem());
                        break;
                    case SLSharedGroupItemsType.Error:
                        sis.Append(ErrorItems[pair.Index].ToErrorItem());
                        break;
                    case SLSharedGroupItemsType.String:
                        sis.Append(StringItems[pair.Index].ToStringItem());
                        break;
                    case SLSharedGroupItemsType.DateTime:
                        sis.Append(DateTimeItems[pair.Index].ToDateTimeItem());
                        break;
                }

            return sis;
        }

        internal SLSharedItems Clone()
        {
            var sis = new SLSharedItems();
            sis.ContainsSemiMixedTypes = ContainsSemiMixedTypes;
            sis.ContainsNonDate = ContainsNonDate;
            sis.ContainsDate = ContainsDate;
            sis.ContainsString = ContainsString;
            sis.ContainsBlank = ContainsBlank;
            sis.ContainsMixedTypes = ContainsMixedTypes;
            sis.ContainsNumber = ContainsNumber;
            sis.ContainsInteger = ContainsInteger;
            sis.MinValue = MinValue;
            sis.MaxValue = MaxValue;
            sis.MinDate = MinDate;
            sis.MaxDate = MaxDate;
            //count
            sis.LongText = LongText;

            sis.Items = new List<SLSharedGroupItemsTypeIndexPair>();
            foreach (var pair in Items)
                sis.Items.Add(new SLSharedGroupItemsTypeIndexPair(pair.Type, pair.Index));

            sis.MissingItems = new List<SLMissingItem>();
            foreach (var mi in MissingItems)
                sis.MissingItems.Add(mi.Clone());

            sis.NumberItems = new List<SLNumberItem>();
            foreach (var ni in NumberItems)
                sis.NumberItems.Add(ni.Clone());

            sis.BooleanItems = new List<SLBooleanItem>();
            foreach (var bi in BooleanItems)
                sis.BooleanItems.Add(bi.Clone());

            sis.ErrorItems = new List<SLErrorItem>();
            foreach (var ei in ErrorItems)
                sis.ErrorItems.Add(ei.Clone());

            sis.StringItems = new List<SLStringItem>();
            foreach (var si in StringItems)
                sis.StringItems.Add(si.Clone());

            sis.DateTimeItems = new List<SLDateTimeItem>();
            foreach (var dti in DateTimeItems)
                sis.DateTimeItems.Add(dti.Clone());

            return sis;
        }
    }
}