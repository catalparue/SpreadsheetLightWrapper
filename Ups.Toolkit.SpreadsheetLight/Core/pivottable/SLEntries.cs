using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.pivottable
{
    internal enum SLEntriesItemsType
    {
        Missing = 0,
        Number,
        Error,
        String
    }

    internal struct SLEntriesItemsTypeIndexPair
    {
        internal SLEntriesItemsType Type;
        // this is the 0-based index into whichever List<> depending on Type.
        internal int Index;

        internal SLEntriesItemsTypeIndexPair(SLEntriesItemsType Type, int Index)
        {
            this.Type = Type;
            this.Index = Index;
        }
    }

    internal class SLEntries
    {
        internal SLEntries()
        {
            SetAllNull();
        }

        internal List<SLEntriesItemsTypeIndexPair> Items { get; set; }

        internal List<SLMissingItem> MissingItems { get; set; }
        internal List<SLNumberItem> NumberItems { get; set; }
        internal List<SLErrorItem> ErrorItems { get; set; }
        internal List<SLStringItem> StringItems { get; set; }

        private void SetAllNull()
        {
            Items = new List<SLEntriesItemsTypeIndexPair>();

            MissingItems = new List<SLMissingItem>();
            NumberItems = new List<SLNumberItem>();
            ErrorItems = new List<SLErrorItem>();
            StringItems = new List<SLStringItem>();
        }

        internal void FromEntries(Entries es)
        {
            SetAllNull();

            SLMissingItem mi;
            SLNumberItem ni;
            SLErrorItem ei;
            SLStringItem si;
            using (var oxr = OpenXmlReader.Create(es))
            {
                while (oxr.Read())
                    if (oxr.ElementType == typeof(MissingItem))
                    {
                        mi = new SLMissingItem();
                        mi.FromMissingItem((MissingItem) oxr.LoadCurrentElement());
                        Items.Add(new SLEntriesItemsTypeIndexPair(SLEntriesItemsType.Missing, MissingItems.Count));
                        MissingItems.Add(mi);
                    }
                    else if (oxr.ElementType == typeof(NumberItem))
                    {
                        ni = new SLNumberItem();
                        ni.FromNumberItem((NumberItem) oxr.LoadCurrentElement());
                        Items.Add(new SLEntriesItemsTypeIndexPair(SLEntriesItemsType.Number, NumberItems.Count));
                        NumberItems.Add(ni);
                    }
                    else if (oxr.ElementType == typeof(ErrorItem))
                    {
                        ei = new SLErrorItem();
                        ei.FromErrorItem((ErrorItem) oxr.LoadCurrentElement());
                        Items.Add(new SLEntriesItemsTypeIndexPair(SLEntriesItemsType.Error, ErrorItems.Count));
                        ErrorItems.Add(ei);
                    }
                    else if (oxr.ElementType == typeof(StringItem))
                    {
                        si = new SLStringItem();
                        si.FromStringItem((StringItem) oxr.LoadCurrentElement());
                        Items.Add(new SLEntriesItemsTypeIndexPair(SLEntriesItemsType.String, StringItems.Count));
                        StringItems.Add(si);
                    }
            }
        }

        internal Entries ToEntries()
        {
            var es = new Entries();

            // is it the sum of all the various items?
            if (Items.Count > 0) es.Count = (uint) Items.Count;

            foreach (var pair in Items)
                switch (pair.Type)
                {
                    case SLEntriesItemsType.Missing:
                        es.Append(MissingItems[pair.Index].ToMissingItem());
                        break;
                    case SLEntriesItemsType.Number:
                        es.Append(NumberItems[pair.Index].ToNumberItem());
                        break;
                    case SLEntriesItemsType.Error:
                        es.Append(ErrorItems[pair.Index].ToErrorItem());
                        break;
                    case SLEntriesItemsType.String:
                        es.Append(StringItems[pair.Index].ToStringItem());
                        break;
                }

            return es;
        }

        internal SLEntries Clone()
        {
            var es = new SLEntries();

            es.Items = new List<SLEntriesItemsTypeIndexPair>();
            foreach (var pair in Items)
                es.Items.Add(new SLEntriesItemsTypeIndexPair(pair.Type, pair.Index));

            es.MissingItems = new List<SLMissingItem>();
            foreach (var mi in MissingItems)
                es.MissingItems.Add(mi.Clone());

            es.NumberItems = new List<SLNumberItem>();
            foreach (var ni in NumberItems)
                es.NumberItems.Add(ni.Clone());

            es.ErrorItems = new List<SLErrorItem>();
            foreach (var ei in ErrorItems)
                es.ErrorItems.Add(ei.Clone());

            es.StringItems = new List<SLStringItem>();
            foreach (var si in StringItems)
                es.StringItems.Add(si.Clone());

            return es;
        }
    }
}