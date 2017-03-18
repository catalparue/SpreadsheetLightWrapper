using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.table
{
    internal class SLDynamicFilter
    {
        internal SLDynamicFilter()
        {
            SetAllNull();
        }

        internal DynamicFilterValues Type { get; set; }
        internal double? Val { get; set; }
        internal double? MaxVal { get; set; }

        private void SetAllNull()
        {
            Type = DynamicFilterValues.Null;
            Val = null;
            MaxVal = null;
        }

        internal void FromDynamicFilter(DynamicFilter df)
        {
            SetAllNull();

            Type = df.Type.Value;
            if (df.Val != null) Val = df.Val.Value;
            if (df.MaxVal != null) MaxVal = df.MaxVal.Value;
        }

        internal DynamicFilter ToDynamicFilter()
        {
            var df = new DynamicFilter();
            df.Type = Type;
            if (Val != null) df.Val = Val.Value;
            if (MaxVal != null) df.MaxVal = MaxVal.Value;

            return df;
        }

        internal SLDynamicFilter Clone()
        {
            var df = new SLDynamicFilter();
            df.Type = Type;
            df.Val = Val;
            df.MaxVal = MaxVal;

            return df;
        }
    }
}