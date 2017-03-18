using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

// Apparently .NET Framework 4 has a System.Tuple, which clashes
// with DocumentFormat.OpenXml.Spreadsheet.Tuple.
// Good thing we're on 3.5...

namespace Ups.Toolkit.SpreadsheetLight.Core.pivottable
{
    /// <summary>
    ///     This doubles for SortByTuple and Tuples
    /// </summary>
    internal class SLTuplesType
    {
        internal SLTuplesType()
        {
            SetAllNull();
        }

        internal List<SLTuple> Tuples { get; set; }
        internal uint? MemberNameCount { get; set; }

        private void SetAllNull()
        {
            Tuples = new List<SLTuple>();
            MemberNameCount = null;
        }

        internal void FromSortByTuple(SortByTuple sbt)
        {
            SetAllNull();

            if (sbt.MemberNameCount != null) MemberNameCount = sbt.MemberNameCount.Value;

            SLTuple t;
            using (var oxr = OpenXmlReader.Create(sbt))
            {
                while (oxr.Read())
                    if (oxr.ElementType == typeof(Tuple))
                    {
                        t = new SLTuple();
                        t.FromTuple((Tuple) oxr.LoadCurrentElement());
                        Tuples.Add(t);
                    }
            }
        }

        internal void FromTuples(Tuples tpls)
        {
            SetAllNull();

            if (tpls.MemberNameCount != null) MemberNameCount = tpls.MemberNameCount.Value;

            SLTuple t;
            using (var oxr = OpenXmlReader.Create(tpls))
            {
                while (oxr.Read())
                    if (oxr.ElementType == typeof(Tuple))
                    {
                        t = new SLTuple();
                        t.FromTuple((Tuple) oxr.LoadCurrentElement());
                        Tuples.Add(t);
                    }
            }
        }

        internal SortByTuple ToSortByTuple()
        {
            var sbt = new SortByTuple();
            if (MemberNameCount != null) sbt.MemberNameCount = MemberNameCount.Value;

            foreach (var t in Tuples)
                sbt.Append(t.ToTuple());

            return sbt;
        }

        internal Tuples ToTuples()
        {
            var tpls = new Tuples();
            if (MemberNameCount != null) tpls.MemberNameCount = MemberNameCount.Value;

            foreach (var t in Tuples)
                tpls.Append(t.ToTuple());

            return tpls;
        }

        internal SLTuplesType Clone()
        {
            var tt = new SLTuplesType();
            tt.MemberNameCount = MemberNameCount;

            tt.Tuples = new List<SLTuple>();
            foreach (var t in Tuples)
                tt.Tuples.Add(t.Clone());

            return tt;
        }
    }
}