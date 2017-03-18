using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.pivottable
{
    internal enum SLPivotCacheRecordItemsType
    {
        Missing = 0,
        Number,
        Boolean,
        Error,
        String,
        DateTime,
        Field
    }

    internal struct SLPivotCacheRecordItemsTypeIndexPair
    {
        internal SLPivotCacheRecordItemsType Type;
        // this is the 0-based index into whichever List<> depending on Type.
        internal int Index;

        internal SLPivotCacheRecordItemsTypeIndexPair(SLPivotCacheRecordItemsType Type, int Index)
        {
            this.Type = Type;
            this.Index = Index;
        }
    }

    internal class SLPivotCacheRecord
    {
        internal SLPivotCacheRecord()
        {
            SetAllNull();
        }

        internal List<SLPivotCacheRecordItemsTypeIndexPair> Items { get; set; }

        internal List<SLMissingItem> MissingItems { get; set; }
        internal List<SLNumberItem> NumberItems { get; set; }
        internal List<SLBooleanItem> BooleanItems { get; set; }
        internal List<SLErrorItem> ErrorItems { get; set; }
        internal List<SLStringItem> StringItems { get; set; }
        internal List<SLDateTimeItem> DateTimeItems { get; set; }
        internal List<uint> FieldItems { get; set; }

        private void SetAllNull()
        {
            Items = new List<SLPivotCacheRecordItemsTypeIndexPair>();

            MissingItems = new List<SLMissingItem>();
            NumberItems = new List<SLNumberItem>();
            BooleanItems = new List<SLBooleanItem>();
            ErrorItems = new List<SLErrorItem>();
            StringItems = new List<SLStringItem>();
            DateTimeItems = new List<SLDateTimeItem>();
            FieldItems = new List<uint>();
        }

        internal void FromPivotCacheRecord(PivotCacheRecord pcr)
        {
            SetAllNull();

            SLMissingItem mi;
            SLNumberItem ni;
            SLBooleanItem bi;
            SLErrorItem ei;
            SLStringItem si;
            SLDateTimeItem dti;
            FieldItem fi;
            using (var oxr = OpenXmlReader.Create(pcr))
            {
                while (oxr.Read())
                    if (oxr.ElementType == typeof(MissingItem))
                    {
                        mi = new SLMissingItem();
                        mi.FromMissingItem((MissingItem) oxr.LoadCurrentElement());
                        Items.Add(new SLPivotCacheRecordItemsTypeIndexPair(SLPivotCacheRecordItemsType.Missing,
                            MissingItems.Count));
                        MissingItems.Add(mi);
                    }
                    else if (oxr.ElementType == typeof(NumberItem))
                    {
                        ni = new SLNumberItem();
                        ni.FromNumberItem((NumberItem) oxr.LoadCurrentElement());
                        Items.Add(new SLPivotCacheRecordItemsTypeIndexPair(SLPivotCacheRecordItemsType.Number,
                            NumberItems.Count));
                        NumberItems.Add(ni);
                    }
                    else if (oxr.ElementType == typeof(BooleanItem))
                    {
                        bi = new SLBooleanItem();
                        bi.FromBooleanItem((BooleanItem) oxr.LoadCurrentElement());
                        Items.Add(new SLPivotCacheRecordItemsTypeIndexPair(SLPivotCacheRecordItemsType.Boolean,
                            BooleanItems.Count));
                        BooleanItems.Add(bi);
                    }
                    else if (oxr.ElementType == typeof(ErrorItem))
                    {
                        ei = new SLErrorItem();
                        ei.FromErrorItem((ErrorItem) oxr.LoadCurrentElement());
                        Items.Add(new SLPivotCacheRecordItemsTypeIndexPair(SLPivotCacheRecordItemsType.Error,
                            ErrorItems.Count));
                        ErrorItems.Add(ei);
                    }
                    else if (oxr.ElementType == typeof(StringItem))
                    {
                        si = new SLStringItem();
                        si.FromStringItem((StringItem) oxr.LoadCurrentElement());
                        Items.Add(new SLPivotCacheRecordItemsTypeIndexPair(SLPivotCacheRecordItemsType.String,
                            StringItems.Count));
                        StringItems.Add(si);
                    }
                    else if (oxr.ElementType == typeof(DateTimeItem))
                    {
                        dti = new SLDateTimeItem();
                        dti.FromDateTimeItem((DateTimeItem) oxr.LoadCurrentElement());
                        Items.Add(new SLPivotCacheRecordItemsTypeIndexPair(SLPivotCacheRecordItemsType.DateTime,
                            DateTimeItems.Count));
                        DateTimeItems.Add(dti);
                    }
                    else if (oxr.ElementType == typeof(FieldItem))
                    {
                        fi = (FieldItem) oxr.LoadCurrentElement();
                        Items.Add(new SLPivotCacheRecordItemsTypeIndexPair(SLPivotCacheRecordItemsType.Field,
                            FieldItems.Count));
                        FieldItems.Add(fi.Val.Value);
                    }
            }
        }

        internal PivotCacheRecord ToPivotCacheRecord()
        {
            var pcr = new PivotCacheRecord();

            foreach (var pair in Items)
                switch (pair.Type)
                {
                    case SLPivotCacheRecordItemsType.Missing:
                        pcr.Append(MissingItems[pair.Index].ToMissingItem());
                        break;
                    case SLPivotCacheRecordItemsType.Number:
                        pcr.Append(NumberItems[pair.Index].ToNumberItem());
                        break;
                    case SLPivotCacheRecordItemsType.Boolean:
                        pcr.Append(BooleanItems[pair.Index].ToBooleanItem());
                        break;
                    case SLPivotCacheRecordItemsType.Error:
                        pcr.Append(ErrorItems[pair.Index].ToErrorItem());
                        break;
                    case SLPivotCacheRecordItemsType.String:
                        pcr.Append(StringItems[pair.Index].ToStringItem());
                        break;
                    case SLPivotCacheRecordItemsType.DateTime:
                        pcr.Append(DateTimeItems[pair.Index].ToDateTimeItem());
                        break;
                    case SLPivotCacheRecordItemsType.Field:
                        pcr.Append(new FieldItem {Val = FieldItems[pair.Index]});
                        break;
                }

            return pcr;
        }

        internal SLPivotCacheRecord Clone()
        {
            var pcr = new SLPivotCacheRecord();

            pcr.Items = new List<SLPivotCacheRecordItemsTypeIndexPair>();
            foreach (var pair in Items)
                pcr.Items.Add(new SLPivotCacheRecordItemsTypeIndexPair(pair.Type, pair.Index));

            pcr.MissingItems = new List<SLMissingItem>();
            foreach (var mi in MissingItems)
                pcr.MissingItems.Add(mi.Clone());

            pcr.NumberItems = new List<SLNumberItem>();
            foreach (var ni in NumberItems)
                pcr.NumberItems.Add(ni.Clone());

            pcr.BooleanItems = new List<SLBooleanItem>();
            foreach (var bi in BooleanItems)
                pcr.BooleanItems.Add(bi.Clone());

            pcr.ErrorItems = new List<SLErrorItem>();
            foreach (var ei in ErrorItems)
                pcr.ErrorItems.Add(ei.Clone());

            pcr.StringItems = new List<SLStringItem>();
            foreach (var si in StringItems)
                pcr.StringItems.Add(si.Clone());

            pcr.DateTimeItems = new List<SLDateTimeItem>();
            foreach (var dti in DateTimeItems)
                pcr.DateTimeItems.Add(dti.Clone());

            pcr.FieldItems = new List<uint>();
            foreach (var i in FieldItems)
                pcr.FieldItems.Add(i);

            return pcr;
        }
    }
}