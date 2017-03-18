using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.pivottable
{
    internal class SLChartFormat
    {
        internal SLChartFormat()
        {
            SetAllNull();
        }

        internal SLPivotArea PivotArea { get; set; }
        internal uint Chart { get; set; }
        internal uint Format { get; set; }
        internal bool Series { get; set; }

        private void SetAllNull()
        {
            PivotArea = new SLPivotArea();
            Chart = 0;
            Format = 0;
            Series = false;
        }

        internal void FromChartFormat(ChartFormat cf)
        {
            SetAllNull();

            if (cf.PivotArea != null) PivotArea.FromPivotArea(cf.PivotArea);

            if (cf.Chart != null) Chart = cf.Chart.Value;
            if (cf.Format != null) Format = cf.Format.Value;
            if (cf.Series != null) Series = cf.Series.Value;
        }

        internal ChartFormat ToChartFormat()
        {
            var cf = new ChartFormat();
            cf.PivotArea = PivotArea.ToPivotArea();

            cf.Chart = Chart;
            cf.Format = Format;
            if (Series) cf.Series = Series;

            return cf;
        }

        internal SLChartFormat Clone()
        {
            var cf = new SLChartFormat();
            cf.PivotArea = PivotArea.Clone();
            cf.Chart = Chart;
            cf.Format = Format;
            cf.Series = Series;

            return cf;
        }
    }
}