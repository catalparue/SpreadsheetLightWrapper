using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.pivottable
{
    internal class SLQuery
    {
        internal bool HasTuples;

        internal SLQuery()
        {
            SetAllNull();
        }

        internal SLTuplesType Tuples { get; set; }

        internal string Mdx { get; set; }

        private void SetAllNull()
        {
            HasTuples = false;
            Tuples = new SLTuplesType();

            Mdx = "";
        }

        internal void FromQuery(Query q)
        {
            SetAllNull();

            if (q.Mdx != null) Mdx = q.Mdx.Value;

            if (q.Tuples != null)
            {
                Tuples.FromTuples(q.Tuples);
                HasTuples = true;
            }
        }

        internal Query ToQuery()
        {
            var q = new Query();
            q.Mdx = Mdx;

            if (HasTuples) q.Tuples = Tuples.ToTuples();

            return q;
        }

        internal SLQuery Clone()
        {
            var q = new SLQuery();
            q.Mdx = Mdx;
            q.HasTuples = HasTuples;
            q.Tuples = Tuples.Clone();

            return q;
        }
    }
}