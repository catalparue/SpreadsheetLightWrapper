using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.pivottable
{
    internal class SLTupleCache
    {
        internal bool HasEntries;

        internal SLTupleCache()
        {
            SetAllNull();
        }

        internal SLEntries Entries { get; set; }
        internal List<SLTupleSet> Sets { get; set; }
        internal List<SLQuery> QueryCache { get; set; }
        internal List<SLServerFormat> ServerFormats { get; set; }

        private void SetAllNull()
        {
            HasEntries = false;
            Entries = new SLEntries();
            Sets = new List<SLTupleSet>();
            QueryCache = new List<SLQuery>();
            ServerFormats = new List<SLServerFormat>();
        }

        internal void FromTupleCache(TupleCache tc)
        {
            SetAllNull();

            // I decided to do this one by one instead of just running through all the child
            // elements. Mainly because this seems safer... so complicated! It's just a pivot table
            // for goodness sakes...

            if (tc.Entries != null)
            {
                Entries.FromEntries(tc.Entries);
                HasEntries = true;
            }

            if (tc.Sets != null)
            {
                SLTupleSet ts;
                using (var oxr = OpenXmlReader.Create(tc.Sets))
                {
                    while (oxr.Read())
                        if (oxr.ElementType == typeof(TupleSet))
                        {
                            ts = new SLTupleSet();
                            ts.FromTupleSet((TupleSet) oxr.LoadCurrentElement());
                            Sets.Add(ts);
                        }
                }
            }

            if (tc.QueryCache != null)
            {
                SLQuery q;
                using (var oxr = OpenXmlReader.Create(tc.QueryCache))
                {
                    while (oxr.Read())
                        if (oxr.ElementType == typeof(Query))
                        {
                            q = new SLQuery();
                            q.FromQuery((Query) oxr.LoadCurrentElement());
                            QueryCache.Add(q);
                        }
                }
            }

            if (tc.ServerFormats != null)
            {
                SLServerFormat sf;
                using (var oxr = OpenXmlReader.Create(tc.ServerFormats))
                {
                    while (oxr.Read())
                        if (oxr.ElementType == typeof(ServerFormat))
                        {
                            sf = new SLServerFormat();
                            sf.FromServerFormat((ServerFormat) oxr.LoadCurrentElement());
                            ServerFormats.Add(sf);
                        }
                }
            }
        }

        internal TupleCache ToTupleCache()
        {
            var tc = new TupleCache();
            if (HasEntries) tc.Entries = Entries.ToEntries();

            if (Sets.Count > 0)
            {
                tc.Sets = new Sets {Count = (uint) Sets.Count};
                foreach (var ts in Sets)
                    tc.Sets.Append(ts.ToTupleSet());
            }

            if (QueryCache.Count > 0)
            {
                tc.QueryCache = new QueryCache {Count = (uint) QueryCache.Count};
                foreach (var q in QueryCache)
                    tc.QueryCache.Append(q.ToQuery());
            }

            if (ServerFormats.Count > 0)
            {
                tc.ServerFormats = new ServerFormats {Count = (uint) ServerFormats.Count};
                foreach (var sf in ServerFormats)
                    tc.ServerFormats.Append(sf.ToServerFormat());
            }

            return tc;
        }

        internal SLTupleCache Clone()
        {
            var tc = new SLTupleCache();
            tc.HasEntries = HasEntries;
            tc.Entries = Entries.Clone();

            tc.Sets = new List<SLTupleSet>();
            foreach (var ts in Sets)
                tc.Sets.Add(ts.Clone());

            tc.QueryCache = new List<SLQuery>();
            foreach (var q in QueryCache)
                tc.QueryCache.Add(q.Clone());

            tc.ServerFormats = new List<SLServerFormat>();
            foreach (var sf in ServerFormats)
                tc.ServerFormats.Add(sf.Clone());

            return tc;
        }
    }
}