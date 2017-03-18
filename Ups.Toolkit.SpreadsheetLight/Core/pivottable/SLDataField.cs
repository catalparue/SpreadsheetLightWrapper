using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.pivottable
{
    internal class SLDataField
    {
        internal SLDataField()
        {
            SetAllNull();
        }

        internal string Name { get; set; }
        internal uint Field { get; set; }
        internal DataConsolidateFunctionValues Subtotal { get; set; }
        internal ShowDataAsValues ShowDataAs { get; set; }
        internal int BaseField { get; set; }
        internal uint BaseItem { get; set; }
        internal uint? NumberFormatId { get; set; }

        private void SetAllNull()
        {
            Name = "";
            Field = 1;
            Subtotal = DataConsolidateFunctionValues.Sum;
            ShowDataAs = ShowDataAsValues.Normal;
            BaseField = -1;

            // why the weird default value? It's 2^20 + 2^8 for what it's worth...
            BaseItem = 1048832;

            NumberFormatId = null;
        }

        internal void FromDataField(DataField df)
        {
            SetAllNull();

            if (df.Name != null) Name = df.Name.Value;
            if (df.Field != null) Field = df.Field.Value;
            if (df.Subtotal != null) Subtotal = df.Subtotal.Value;
            if (df.ShowDataAs != null) ShowDataAs = df.ShowDataAs.Value;
            if (df.BaseField != null) BaseField = df.BaseField.Value;
            if (df.BaseItem != null) BaseItem = df.BaseItem.Value;
            if (df.NumberFormatId != null) NumberFormatId = df.NumberFormatId.Value;
        }

        internal DataField ToDataField()
        {
            var df = new DataField();
            if ((Name != null) && (Name.Length > 0)) df.Name = Name;
            df.Field = Field;
            if (Subtotal != DataConsolidateFunctionValues.Sum) df.Subtotal = Subtotal;
            if (ShowDataAs != ShowDataAsValues.Normal) df.ShowDataAs = ShowDataAs;
            if (BaseField != -1) df.BaseField = BaseField;
            if (BaseItem != 1048832) df.BaseItem = BaseItem;
            if (NumberFormatId != null) df.NumberFormatId = NumberFormatId.Value;

            return df;
        }

        internal SLDataField Clone()
        {
            var df = new SLDataField();
            df.Name = Name;
            df.Field = Field;
            df.Subtotal = Subtotal;
            df.ShowDataAs = ShowDataAs;
            df.BaseField = BaseField;
            df.BaseItem = BaseItem;
            df.NumberFormatId = NumberFormatId;

            return df;
        }
    }
}