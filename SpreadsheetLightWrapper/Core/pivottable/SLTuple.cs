using DocumentFormat.OpenXml.Spreadsheet;

// Apparently .NET Framework 4 has a System.Tuple, which clashes
// with DocumentFormat.OpenXml.Spreadsheet.Tuple.
// Good thing we're on 3.5...

namespace SpreadsheetLightWrapper.Core.pivottable
{
    internal class SLTuple
    {
        internal SLTuple()
        {
            SetAllNull();
        }

        internal uint? Field { get; set; }
        internal uint? Hierarchy { get; set; }
        internal uint Item { get; set; }

        private void SetAllNull()
        {
            Field = null;
            Hierarchy = null;
            Item = 0;
        }

        internal void FromTuple(Tuple t)
        {
            SetAllNull();

            if (t.Field != null) Field = t.Field.Value;
            if (t.Hierarchy != null) Hierarchy = t.Hierarchy.Value;
            if (t.Item != null) Item = t.Item.Value;
        }

        internal Tuple ToTuple()
        {
            var t = new Tuple();
            if (Field != null) t.Field = Field.Value;
            if (Hierarchy != null) t.Hierarchy = Hierarchy.Value;
            t.Item = Item;

            return t;
        }

        internal SLTuple Clone()
        {
            var t = new SLTuple();
            t.Field = Field;
            t.Hierarchy = Hierarchy;
            t.Item = Item;

            return t;
        }
    }
}