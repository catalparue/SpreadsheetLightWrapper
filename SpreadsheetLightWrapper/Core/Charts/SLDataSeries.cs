using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLightWrapper.Core.Charts
{
    internal class SLDataSeries
    {
        // this is SeriesText
        internal bool? IsStringReference;

        internal SLDataSeries(List<Color> ThemeColors)
        {
            ChartType = SLDataSeriesChartType.None;

            Index = 0;
            Order = 0;

            IsStringReference = null;
            StringReference = new SLStringReference();
            NumericValue = string.Empty;

            Options = new SLDataSeriesOptions(ThemeColors);

            DataPointOptionsList = new Dictionary<int, SLDataPointOptions>();

            GroupDataLabelOptions = null;
            DataLabelOptionsList = new Dictionary<int, SLDataLabelOptions>();

            BubbleSize = new SLNumberDataSourceType();

            AxisData = new SLAxisDataSourceType();
            NumberData = new SLNumberDataSourceType();
        }

        internal SLDataSeriesChartType ChartType { get; set; }

        internal uint Index { get; set; }
        internal uint Order { get; set; }
        internal SLStringReference StringReference { get; set; }
        internal string NumericValue { get; set; }

        internal SLDataSeriesOptions Options { get; set; }

        //PictureOptions

        internal Dictionary<int, SLDataPointOptions> DataPointOptionsList { get; set; }

        internal SLGroupDataLabelOptions GroupDataLabelOptions { get; set; }
        internal Dictionary<int, SLDataLabelOptions> DataLabelOptionsList { get; set; }

        //List<Trendline>
        //List<ErrorBars>

        //category
        //value

        //xval
        //yval

        internal SLNumberDataSourceType BubbleSize { get; set; }

        internal SLAxisDataSourceType AxisData { get; set; }
        internal SLNumberDataSourceType NumberData { get; set; }

        internal C.PieChartSeries ToPieChartSeries(bool IsStylish = false)
        {
            var pcs = new C.PieChartSeries();
            pcs.Index = new C.Index {Val = Index};
            pcs.Order = new C.Order {Val = Order};

            if (IsStringReference != null)
            {
                pcs.SeriesText = new C.SeriesText();
                if (IsStringReference.Value)
                    pcs.SeriesText.StringReference = StringReference.ToStringReference();
                else
                    pcs.SeriesText.NumericValue = new C.NumericValue(NumericValue);
            }

            if (Options.ShapeProperties.HasShapeProperties)
                pcs.ChartShapeProperties = Options.ShapeProperties.ToChartShapeProperties(IsStylish);

            if (Options.iExplosion != null)
                pcs.Explosion = new C.Explosion {Val = Options.Explosion};

            if (DataPointOptionsList.Count > 0)
            {
                var indexlist = DataPointOptionsList.Keys.ToList();
                indexlist.Sort();
                int index;
                for (var i = 0; i < indexlist.Count; ++i)
                {
                    index = indexlist[i];
                    pcs.Append(DataPointOptionsList[index].ToDataPoint(index, IsStylish));
                }
            }

            if ((GroupDataLabelOptions != null) || (DataLabelOptionsList.Count > 0))
                if (GroupDataLabelOptions == null)
                {
                    var gdloptions = new SLGroupDataLabelOptions(new List<Color>());
                    pcs.Append(gdloptions.ToDataLabels(DataLabelOptionsList, true));
                }
                else
                {
                    pcs.Append(GroupDataLabelOptions.ToDataLabels(DataLabelOptionsList, false));
                }

            pcs.Append(AxisData.ToCategoryAxisData());
            pcs.Append(NumberData.ToValues());

            return pcs;
        }

        internal C.RadarChartSeries ToRadarChartSeries(bool IsStylish = false)
        {
            var rcs = new C.RadarChartSeries();
            rcs.Index = new C.Index {Val = Index};
            rcs.Order = new C.Order {Val = Order};

            if (IsStringReference != null)
            {
                rcs.SeriesText = new C.SeriesText();
                if (IsStringReference.Value)
                    rcs.SeriesText.StringReference = StringReference.ToStringReference();
                else
                    rcs.SeriesText.NumericValue = new C.NumericValue(NumericValue);
            }

            if (Options.ShapeProperties.HasShapeProperties)
                rcs.ChartShapeProperties = Options.ShapeProperties.ToChartShapeProperties(IsStylish);

            if (Options.Marker.HasMarker)
                rcs.Marker = Options.Marker.ToMarker(IsStylish);

            if (DataPointOptionsList.Count > 0)
            {
                var indexlist = DataPointOptionsList.Keys.ToList();
                indexlist.Sort();
                int index;
                for (var i = 0; i < indexlist.Count; ++i)
                {
                    index = indexlist[i];
                    rcs.Append(DataPointOptionsList[index].ToDataPoint(index, IsStylish));
                }
            }

            if ((GroupDataLabelOptions != null) || (DataLabelOptionsList.Count > 0))
                if (GroupDataLabelOptions == null)
                {
                    var gdloptions = new SLGroupDataLabelOptions(new List<Color>());
                    rcs.Append(gdloptions.ToDataLabels(DataLabelOptionsList, true));
                }
                else
                {
                    rcs.Append(GroupDataLabelOptions.ToDataLabels(DataLabelOptionsList, false));
                }

            rcs.Append(AxisData.ToCategoryAxisData());
            rcs.Append(NumberData.ToValues());

            return rcs;
        }

        internal C.AreaChartSeries ToAreaChartSeries(bool IsStylish = false)
        {
            var acs = new C.AreaChartSeries();
            acs.Index = new C.Index {Val = Index};
            acs.Order = new C.Order {Val = Order};

            if (IsStringReference != null)
            {
                acs.SeriesText = new C.SeriesText();
                if (IsStringReference.Value)
                    acs.SeriesText.StringReference = StringReference.ToStringReference();
                else
                    acs.SeriesText.NumericValue = new C.NumericValue(NumericValue);
            }

            if (Options.ShapeProperties.HasShapeProperties)
                acs.ChartShapeProperties = Options.ShapeProperties.ToChartShapeProperties(IsStylish);

            //PictureOptions

            if (DataPointOptionsList.Count > 0)
            {
                var indexlist = DataPointOptionsList.Keys.ToList();
                indexlist.Sort();
                int index;
                for (var i = 0; i < indexlist.Count; ++i)
                {
                    index = indexlist[i];
                    acs.Append(DataPointOptionsList[index].ToDataPoint(index, IsStylish));
                }
            }

            if ((GroupDataLabelOptions != null) || (DataLabelOptionsList.Count > 0))
                if (GroupDataLabelOptions == null)
                {
                    var gdloptions = new SLGroupDataLabelOptions(new List<Color>());
                    acs.Append(gdloptions.ToDataLabels(DataLabelOptionsList, true));
                }
                else
                {
                    acs.Append(GroupDataLabelOptions.ToDataLabels(DataLabelOptionsList, false));
                }

            acs.Append(AxisData.ToCategoryAxisData());
            acs.Append(NumberData.ToValues());

            return acs;
        }

        internal C.BarChartSeries ToBarChartSeries(bool IsStylish = false)
        {
            var bcs = new C.BarChartSeries();
            bcs.Index = new C.Index {Val = Index};
            bcs.Order = new C.Order {Val = Order};

            if (IsStringReference != null)
            {
                bcs.SeriesText = new C.SeriesText();
                if (IsStringReference.Value)
                    bcs.SeriesText.StringReference = StringReference.ToStringReference();
                else
                    bcs.SeriesText.NumericValue = new C.NumericValue(NumericValue);
            }

            if (Options.ShapeProperties.HasShapeProperties)
                bcs.ChartShapeProperties = Options.ShapeProperties.ToChartShapeProperties(IsStylish);

            bcs.InvertIfNegative = new C.InvertIfNegative {Val = Options.InvertIfNegative ?? false};

            //PictureOptions

            if (DataPointOptionsList.Count > 0)
            {
                var indexlist = DataPointOptionsList.Keys.ToList();
                indexlist.Sort();
                int index;
                for (var i = 0; i < indexlist.Count; ++i)
                {
                    index = indexlist[i];
                    bcs.Append(DataPointOptionsList[index].ToDataPoint(index, IsStylish));
                }
            }

            if ((GroupDataLabelOptions != null) || (DataLabelOptionsList.Count > 0))
                if (GroupDataLabelOptions == null)
                {
                    var gdloptions = new SLGroupDataLabelOptions(new List<Color>());
                    bcs.Append(gdloptions.ToDataLabels(DataLabelOptionsList, true));
                }
                else
                {
                    bcs.Append(GroupDataLabelOptions.ToDataLabels(DataLabelOptionsList, false));
                }

            bcs.Append(AxisData.ToCategoryAxisData());
            bcs.Append(NumberData.ToValues());

            if (Options.vShape != null)
                bcs.Append(new C.Shape {Val = Options.vShape.Value});

            return bcs;
        }

        internal C.ScatterChartSeries ToScatterChartSeries(bool IsStylish = false)
        {
            var scs = new C.ScatterChartSeries();
            scs.Index = new C.Index {Val = Index};
            scs.Order = new C.Order {Val = Order};

            if (IsStringReference != null)
            {
                scs.SeriesText = new C.SeriesText();
                if (IsStringReference.Value)
                    scs.SeriesText.StringReference = StringReference.ToStringReference();
                else
                    scs.SeriesText.NumericValue = new C.NumericValue(NumericValue);
            }

            if (Options.ShapeProperties.HasShapeProperties)
                scs.ChartShapeProperties = Options.ShapeProperties.ToChartShapeProperties(IsStylish);

            if (Options.Marker.HasMarker)
                scs.Marker = Options.Marker.ToMarker(IsStylish);

            if (DataPointOptionsList.Count > 0)
            {
                var indexlist = DataPointOptionsList.Keys.ToList();
                indexlist.Sort();
                int index;
                for (var i = 0; i < indexlist.Count; ++i)
                {
                    index = indexlist[i];
                    scs.Append(DataPointOptionsList[index].ToDataPoint(index, IsStylish));
                }
            }

            if ((GroupDataLabelOptions != null) || (DataLabelOptionsList.Count > 0))
                if (GroupDataLabelOptions == null)
                {
                    var gdloptions = new SLGroupDataLabelOptions(new List<Color>());
                    scs.Append(gdloptions.ToDataLabels(DataLabelOptionsList, true));
                }
                else
                {
                    scs.Append(GroupDataLabelOptions.ToDataLabels(DataLabelOptionsList, false));
                }

            scs.Append(AxisData.ToXValues());
            scs.Append(NumberData.ToYValues());

            scs.Append(new C.Smooth {Val = Options.Smooth});

            return scs;
        }

        internal C.LineChartSeries ToLineChartSeries(bool IsStylish = false)
        {
            var lcs = new C.LineChartSeries();
            lcs.Index = new C.Index {Val = Index};
            lcs.Order = new C.Order {Val = Order};

            if (IsStringReference != null)
            {
                lcs.SeriesText = new C.SeriesText();
                if (IsStringReference.Value)
                    lcs.SeriesText.StringReference = StringReference.ToStringReference();
                else
                    lcs.SeriesText.NumericValue = new C.NumericValue(NumericValue);
            }

            if (Options.ShapeProperties.HasShapeProperties)
                lcs.ChartShapeProperties = Options.ShapeProperties.ToChartShapeProperties(IsStylish);

            if (Options.Marker.HasMarker)
                lcs.Marker = Options.Marker.ToMarker(IsStylish);

            if (DataPointOptionsList.Count > 0)
            {
                var indexlist = DataPointOptionsList.Keys.ToList();
                indexlist.Sort();
                int index;
                for (var i = 0; i < indexlist.Count; ++i)
                {
                    index = indexlist[i];
                    lcs.Append(DataPointOptionsList[index].ToDataPoint(index, IsStylish));
                }
            }

            if ((GroupDataLabelOptions != null) || (DataLabelOptionsList.Count > 0))
                if (GroupDataLabelOptions == null)
                {
                    var gdloptions = new SLGroupDataLabelOptions(new List<Color>());
                    lcs.Append(gdloptions.ToDataLabels(DataLabelOptionsList, true));
                }
                else
                {
                    lcs.Append(GroupDataLabelOptions.ToDataLabels(DataLabelOptionsList, false));
                }

            lcs.Append(AxisData.ToCategoryAxisData());
            lcs.Append(NumberData.ToValues());

            lcs.Append(new C.Smooth {Val = Options.Smooth});

            return lcs;
        }

        internal C.BubbleChartSeries ToBubbleChartSeries(bool IsStylish = false)
        {
            var bcs = new C.BubbleChartSeries();
            bcs.Index = new C.Index {Val = Index};
            bcs.Order = new C.Order {Val = Order};

            if (IsStringReference != null)
            {
                bcs.SeriesText = new C.SeriesText();
                if (IsStringReference.Value)
                    bcs.SeriesText.StringReference = StringReference.ToStringReference();
                else
                    bcs.SeriesText.NumericValue = new C.NumericValue(NumericValue);
            }

            if (Options.ShapeProperties.HasShapeProperties)
                bcs.ChartShapeProperties = Options.ShapeProperties.ToChartShapeProperties(IsStylish);

            bcs.InvertIfNegative = new C.InvertIfNegative {Val = Options.InvertIfNegative ?? false};

            if (DataPointOptionsList.Count > 0)
            {
                var indexlist = DataPointOptionsList.Keys.ToList();
                indexlist.Sort();
                int index;
                for (var i = 0; i < indexlist.Count; ++i)
                {
                    index = indexlist[i];
                    bcs.Append(DataPointOptionsList[index].ToDataPoint(index, IsStylish));
                }
            }

            if ((GroupDataLabelOptions != null) || (DataLabelOptionsList.Count > 0))
                if (GroupDataLabelOptions == null)
                {
                    var gdloptions = new SLGroupDataLabelOptions(new List<Color>());
                    bcs.Append(gdloptions.ToDataLabels(DataLabelOptionsList, true));
                }
                else
                {
                    bcs.Append(GroupDataLabelOptions.ToDataLabels(DataLabelOptionsList, false));
                }

            bcs.Append(AxisData.ToXValues());
            bcs.Append(NumberData.ToYValues());
            bcs.Append(BubbleSize.ToBubbleSize());

            if (Options.bBubble3D != null)
                bcs.Append(new C.Bubble3D {Val = Options.Bubble3D});

            return bcs;
        }

        internal C.SurfaceChartSeries ToSurfaceChartSeries(bool IsStylish = false)
        {
            var scs = new C.SurfaceChartSeries();
            scs.Index = new C.Index {Val = Index};
            scs.Order = new C.Order {Val = Order};

            if (IsStringReference != null)
            {
                scs.SeriesText = new C.SeriesText();
                if (IsStringReference.Value)
                    scs.SeriesText.StringReference = StringReference.ToStringReference();
                else
                    scs.SeriesText.NumericValue = new C.NumericValue(NumericValue);
            }

            if (Options.ShapeProperties.HasShapeProperties)
                scs.ChartShapeProperties = Options.ShapeProperties.ToChartShapeProperties(IsStylish);

            scs.Append(AxisData.ToCategoryAxisData());
            scs.Append(NumberData.ToValues());

            return scs;
        }

        internal SLDataSeries Clone()
        {
            var ds = new SLDataSeries(Options.ShapeProperties.listThemeColors);
            ds.ChartType = ChartType;
            ds.Index = Index;
            ds.Order = Order;
            ds.IsStringReference = IsStringReference;
            ds.StringReference = StringReference.Clone();
            ds.NumericValue = NumericValue;
            ds.Options = Options.Clone();

            var keys = DataPointOptionsList.Keys.ToList();
            ds.DataPointOptionsList = new Dictionary<int, SLDataPointOptions>();
            foreach (var index in keys)
                ds.DataPointOptionsList[index] = DataPointOptionsList[index].Clone();

            if (GroupDataLabelOptions != null) ds.GroupDataLabelOptions = GroupDataLabelOptions.Clone();

            keys = DataLabelOptionsList.Keys.ToList();
            ds.DataLabelOptionsList = new Dictionary<int, SLDataLabelOptions>();
            foreach (var index in keys)
                ds.DataLabelOptionsList[index] = DataLabelOptionsList[index].Clone();

            ds.BubbleSize = BubbleSize.Clone();
            ds.AxisData = AxisData.Clone();
            ds.NumberData = NumberData.Clone();

            return ds;
        }
    }
}