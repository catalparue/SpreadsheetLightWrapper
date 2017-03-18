using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.pivottable
{
    internal class SLPageField
    {
        internal SLPageField()
        {
            SetAllNull();
        }

        internal int Field { get; set; }
        internal uint? Item { get; set; }
        internal int Hierarchy { get; set; }
        internal string Name { get; set; }
        internal string Caption { get; set; }

        private void SetAllNull()
        {
            Field = 0;
            Item = null;
            Hierarchy = 0;
            Name = "";
            Caption = "";
        }

        internal void FromPageField(PageField pf)
        {
            SetAllNull();

            if (pf.Field != null) Field = pf.Field.Value;
            if (pf.Item != null) Item = pf.Item.Value;
            if (pf.Hierarchy != null) Hierarchy = pf.Hierarchy.Value;
            if (pf.Name != null) Name = pf.Name.Value;
            if (pf.Caption != null) Caption = pf.Caption.Value;
        }

        internal PageField ToPageField()
        {
            var pf = new PageField();
            pf.Field = Field;
            if (Item != null) pf.Item = Item.Value;
            pf.Hierarchy = Hierarchy;
            if ((Name != null) && (Name.Length > 0)) pf.Name = Name;
            if ((Caption != null) && (Caption.Length > 0)) pf.Caption = Caption;

            return pf;
        }

        internal SLPageField Clone()
        {
            var pf = new SLPageField();
            pf.Field = Field;
            pf.Item = Item;
            pf.Hierarchy = Hierarchy;
            pf.Name = Name;
            pf.Caption = Caption;

            return pf;
        }
    }
}