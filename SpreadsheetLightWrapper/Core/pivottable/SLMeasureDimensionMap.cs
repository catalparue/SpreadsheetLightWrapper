using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.pivottable
{
    internal class SLMeasureDimensionMap
    {
        internal SLMeasureDimensionMap()
        {
            SetAllNull();
        }

        internal uint? MeasureGroup { get; set; }
        internal uint? Dimension { get; set; }

        private void SetAllNull()
        {
            MeasureGroup = null;
            Dimension = null;
        }

        internal void FromMeasureDimensionMap(MeasureDimensionMap mdm)
        {
            SetAllNull();

            if (mdm.MeasureGroup != null) MeasureGroup = mdm.MeasureGroup.Value;
            if (mdm.Dimension != null) Dimension = mdm.Dimension.Value;
        }

        internal MeasureDimensionMap ToMeasureDimensionMap()
        {
            var mdm = new MeasureDimensionMap();
            if (MeasureGroup != null) mdm.MeasureGroup = MeasureGroup.Value;
            if (Dimension != null) mdm.Dimension = Dimension.Value;

            return mdm;
        }

        internal SLMeasureDimensionMap Clone()
        {
            var mdm = new SLMeasureDimensionMap();
            mdm.MeasureGroup = MeasureGroup;
            mdm.Dimension = Dimension;

            return mdm;
        }
    }
}