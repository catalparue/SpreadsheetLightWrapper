using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.pivottable
{
    internal class SLDimension
    {
        internal SLDimension()
        {
            SetAllNull();
        }

        internal bool Measure { get; set; }
        internal string Name { get; set; }
        internal string UniqueName { get; set; }
        internal string Caption { get; set; }

        private void SetAllNull()
        {
            Measure = false;
            Name = "";
            UniqueName = "";
            Caption = "";
        }

        internal void FromDimension(Dimension d)
        {
            SetAllNull();

            if (d.Measure != null) Measure = d.Measure.Value;
            if (d.Name != null) Name = d.Name.Value;
            if (d.UniqueName != null) UniqueName = d.UniqueName.Value;
            if (d.Caption != null) Caption = d.Caption.Value;
        }

        internal Dimension ToDimension()
        {
            var d = new Dimension();
            if (Measure) d.Measure = Measure;
            d.Name = Name;
            d.UniqueName = UniqueName;
            d.Caption = Caption;

            return d;
        }

        internal SLDimension Clone()
        {
            var d = new SLDimension();
            d.Measure = Measure;
            d.Name = Name;
            d.UniqueName = UniqueName;
            d.Caption = Caption;

            return d;
        }
    }
}