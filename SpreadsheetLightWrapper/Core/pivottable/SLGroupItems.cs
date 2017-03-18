using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.pivottable
{
    internal class SLGroupItems
    {
        internal SLGroupItems()
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

        private void SetAllNull()
        {
            Items = new List<SLSharedGroupItemsTypeIndexPair>();

            MissingItems = new List<SLMissingItem>();
            NumberItems = new List<SLNumberItem>();
            BooleanItems = new List<SLBooleanItem>();
            ErrorItems = new List<SLErrorItem>();
            StringItems = new List<SLStringItem>();
            DateTimeItems = new List<SLDateTimeItem>();
        }

        internal void FromGroupItems(GroupItems gis)
        {
            SetAllNull();

            SLMissingItem mi;
            SLNumberItem ni;
            SLBooleanItem bi;
            SLErrorItem ei;
            SLStringItem si;
            SLDateTimeItem dti;
            using (var oxr = OpenXmlReader.Create(gis))
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

        internal GroupItems ToGroupItems()
        {
            var gis = new GroupItems();
            gis.Count = (uint) Items.Count;

            foreach (var pair in Items)
                switch (pair.Type)
                {
                    case SLSharedGroupItemsType.Missing:
                        gis.Append(MissingItems[pair.Index].ToMissingItem());
                        break;
                    case SLSharedGroupItemsType.Number:
                        gis.Append(NumberItems[pair.Index].ToNumberItem());
                        break;
                    case SLSharedGroupItemsType.Boolean:
                        gis.Append(BooleanItems[pair.Index].ToBooleanItem());
                        break;
                    case SLSharedGroupItemsType.Error:
                        gis.Append(ErrorItems[pair.Index].ToErrorItem());
                        break;
                    case SLSharedGroupItemsType.String:
                        gis.Append(StringItems[pair.Index].ToStringItem());
                        break;
                    case SLSharedGroupItemsType.DateTime:
                        gis.Append(DateTimeItems[pair.Index].ToDateTimeItem());
                        break;
                }

            return gis;
        }

        internal SLGroupItems Clone()
        {
            var gis = new SLGroupItems();

            gis.Items = new List<SLSharedGroupItemsTypeIndexPair>();
            foreach (var pair in Items)
                gis.Items.Add(new SLSharedGroupItemsTypeIndexPair(pair.Type, pair.Index));

            gis.MissingItems = new List<SLMissingItem>();
            foreach (var mi in MissingItems)
                gis.MissingItems.Add(mi.Clone());

            gis.NumberItems = new List<SLNumberItem>();
            foreach (var ni in NumberItems)
                gis.NumberItems.Add(ni.Clone());

            gis.BooleanItems = new List<SLBooleanItem>();
            foreach (var bi in BooleanItems)
                gis.BooleanItems.Add(bi.Clone());

            gis.ErrorItems = new List<SLErrorItem>();
            foreach (var ei in ErrorItems)
                gis.ErrorItems.Add(ei.Clone());

            gis.StringItems = new List<SLStringItem>();
            foreach (var si in StringItems)
                gis.StringItems.Add(si.Clone());

            gis.DateTimeItems = new List<SLDateTimeItem>();
            foreach (var dti in DateTimeItems)
                gis.DateTimeItems.Add(dti.Clone());

            return gis;
        }
    }
}