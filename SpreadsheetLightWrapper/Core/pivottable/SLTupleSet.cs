using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.pivottable
{
    internal class SLTupleSet
    {
        internal bool HasSortByTuple;

        internal SLTupleSet()
        {
            SetAllNull();
        }

        //CT_Set

        internal List<SLTuplesType> Tuples { get; set; }
        internal SLTuplesType SortByTuple { get; set; }

        // count is for number of Tuples
        internal int MaxRank { get; set; }
        internal string SetDefinition { get; set; }
        internal SortValues SortType { get; set; }
        internal bool QueryFailed { get; set; }

        private void SetAllNull()
        {
            Tuples = new List<SLTuplesType>();

            HasSortByTuple = false;
            SortByTuple = new SLTuplesType();

            MaxRank = 0;
            SetDefinition = "";
            SortType = SortValues.None;
            QueryFailed = false;
        }

        internal void FromTupleSet(TupleSet ts)
        {
            SetAllNull();

            if (ts.MaxRank != null) MaxRank = ts.MaxRank.Value;
            if (ts.SetDefinition != null) SetDefinition = ts.SetDefinition.Value;
            if (ts.SortType != null) SortType = ts.SortType.Value;
            if (ts.QueryFailed != null) QueryFailed = ts.QueryFailed.Value;

            SLTuplesType tt;
            using (var oxr = OpenXmlReader.Create(ts))
            {
                while (oxr.Read())
                    if (oxr.ElementType == typeof(Tuples))
                    {
                        tt = new SLTuplesType();
                        tt.FromTuples((Tuples) oxr.LoadCurrentElement());
                        Tuples.Add(tt);
                    }
                    else if (oxr.ElementType == typeof(SortByTuple))
                    {
                        SortByTuple.FromSortByTuple((SortByTuple) oxr.LoadCurrentElement());
                        HasSortByTuple = true;
                    }
            }
        }

        internal TupleSet ToTupleSet()
        {
            var ts = new TupleSet();
            if (Tuples.Count > 0) ts.Count = (uint) Tuples.Count;
            ts.MaxRank = MaxRank;
            ts.SetDefinition = SetDefinition;
            if (SortType != SortValues.None) ts.SortType = SortType;
            if (QueryFailed) ts.QueryFailed = QueryFailed;

            if (Tuples.Count > 0)
                foreach (var tt in Tuples)
                    ts.Append(tt.ToTuples());

            if (HasSortByTuple)
                ts.Append(SortByTuple.ToSortByTuple());

            return ts;
        }

        internal SLTupleSet Clone()
        {
            var ts = new SLTupleSet();
            ts.MaxRank = MaxRank;
            ts.SetDefinition = SetDefinition;
            ts.SortType = SortType;
            ts.QueryFailed = QueryFailed;

            ts.Tuples = new List<SLTuplesType>();
            foreach (var tt in Tuples)
                ts.Tuples.Add(tt.Clone());

            ts.HasSortByTuple = HasSortByTuple;
            ts.SortByTuple = SortByTuple.Clone();

            return ts;
        }
    }
}