using System.Collections.Generic;
using System.Drawing;
using SpreadsheetLightWrapper.Core.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace SpreadsheetLightWrapper.Core.Charts
{
    internal class SLChartOptions
    {
        internal bool? bWireframe;

        private byte byHoleSize;

        private sbyte byOverlap;

        internal bool HasDropLines;

        internal bool HasHighLowLines;

        internal bool HasSplit;

        internal bool HasUpDownBars;

        private uint iBubbleScale;

        private ushort iFirstSliceAngle;

        private ushort iGapDepth;

        private ushort iGapWidth;

        private ushort iSecondPieSize;

        // for the series line of of-pie charts
        internal SLShapeProperties SeriesLinesShapeProperties;

        internal SLChartOptions(List<Color> ThemeColors, bool IsStylish = false)
        {
            BarDirection = C.BarDirectionValues.Bar;
            BarGrouping = C.BarGroupingValues.Standard;
            VaryColors = null;
            GapWidth = 150;
            GapDepth = 150;
            Overlap = 0;
            Shape = C.ShapeValues.Box;
            Grouping = C.GroupingValues.Standard;
            ShowMarker = true;
            Smooth = false;
            FirstSliceAngle = 0;
            HoleSize = 10;
            HasSplit = false;
            SplitType = C.SplitValues.Position;
            SplitPosition = 0;
            SecondPiePoints = new List<int>();
            SecondPieSize = 75;
            SeriesLinesShapeProperties = new SLShapeProperties(ThemeColors);
            ScatterStyle = C.ScatterStyleValues.Line;
            bWireframe = null;
            RadarStyle = C.RadarStyleValues.Standard;
            Bubble3D = true;
            BubbleScale = 100;
            ShowNegativeBubbles = true;
            SizeRepresents = C.SizeRepresentsValues.Area;
            HasDropLines = false;
            DropLines = new SLDropLines(ThemeColors, IsStylish);
            HasHighLowLines = false;
            HighLowLines = new SLHighLowLines(ThemeColors, IsStylish);
            HasUpDownBars = false;
            UpDownBars = new SLUpDownBars(ThemeColors, IsStylish);
        }

        internal C.BarDirectionValues BarDirection { get; set; }
        internal C.BarGroupingValues BarGrouping { get; set; }

        internal bool? VaryColors { get; set; }

        internal ushort GapWidth
        {
            get { return iGapWidth; }
            set
            {
                iGapWidth = value;
                if (iGapWidth > 500) iGapWidth = 500;
            }
        }

        internal ushort GapDepth
        {
            get { return iGapDepth; }
            set
            {
                iGapDepth = value;
                if (iGapDepth > 500) iGapDepth = 500;
            }
        }

        internal sbyte Overlap
        {
            get { return byOverlap; }
            set
            {
                byOverlap = value;
                if (byOverlap < -100) byOverlap = -100;
                if (byOverlap > 100) byOverlap = 100;
            }
        }

        internal C.ShapeValues Shape { get; set; }

        internal C.GroupingValues Grouping { get; set; }

        internal bool ShowMarker { get; set; }
        internal bool Smooth { get; set; }

        internal ushort FirstSliceAngle
        {
            get { return iFirstSliceAngle; }
            set
            {
                iFirstSliceAngle = value;
                if (iFirstSliceAngle > 360) iFirstSliceAngle = 360;
            }
        }

        internal byte HoleSize
        {
            get { return byHoleSize; }
            set
            {
                byHoleSize = value;
                if (byHoleSize < 10) byHoleSize = 10;
                if (byHoleSize > 90) byHoleSize = 90;
            }
        }

        internal C.SplitValues SplitType { get; set; }
        internal double SplitPosition { get; set; }
        internal List<int> SecondPiePoints { get; set; }

        internal ushort SecondPieSize
        {
            get { return iSecondPieSize; }
            set
            {
                iSecondPieSize = value;
                if (iSecondPieSize < 5) iSecondPieSize = 5;
                if (iSecondPieSize > 200) iSecondPieSize = 200;
            }
        }

        internal C.ScatterStyleValues ScatterStyle { get; set; }

        internal bool Wireframe
        {
            get { return bWireframe ?? true; }
            set { bWireframe = value; }
        }

        internal C.RadarStyleValues RadarStyle { get; set; }

        internal bool Bubble3D { get; set; }

        internal uint BubbleScale
        {
            get { return iBubbleScale; }
            set
            {
                iBubbleScale = value;
                if (iBubbleScale > 300) iBubbleScale = 300;
            }
        }

        internal bool ShowNegativeBubbles { get; set; }
        internal C.SizeRepresentsValues SizeRepresents { get; set; }
        internal SLDropLines DropLines { get; set; }
        internal SLHighLowLines HighLowLines { get; set; }
        internal SLUpDownBars UpDownBars { get; set; }

        internal void MergeOptions(SLBarChartOptions bco)
        {
            GapWidth = bco.GapWidth;
            GapDepth = bco.GapDepth;
            Overlap = bco.Overlap;
        }

        internal void MergeOptions(SLLineChartOptions lco)
        {
            GapDepth = lco.GapDepth;
            HasDropLines = lco.HasDropLines;
            DropLines = lco.DropLines.Clone();
            HasHighLowLines = lco.HasHighLowLines;
            HighLowLines = lco.HighLowLines.Clone();
            HasUpDownBars = lco.HasUpDownBars;
            UpDownBars = lco.UpDownBars.Clone();
            Smooth = lco.Smooth;
        }

        internal void MergeOptions(SLPieChartOptions pco)
        {
            VaryColors = pco.VaryColors;
            FirstSliceAngle = pco.FirstSliceAngle;
            HoleSize = pco.HoleSize;
            GapWidth = pco.GapWidth;
            HasSplit = pco.HasSplit;
            SplitType = pco.SplitType;
            SplitPosition = pco.SplitPosition;

            SecondPiePoints.Clear();
            foreach (var i in pco.SecondPiePoints)
                SecondPiePoints.Add(i);
            SecondPiePoints.Sort();

            SecondPieSize = pco.SecondPieSize;

            SeriesLinesShapeProperties = pco.ShapeProperties.Clone();
        }

        internal void MergeOptions(SLAreaChartOptions aco)
        {
            HasDropLines = aco.HasDropLines;
            DropLines = aco.DropLines.Clone();
            GapDepth = aco.GapDepth;
        }

        internal void MergeOptions(SLBubbleChartOptions bco)
        {
            Bubble3D = bco.Bubble3D;
            BubbleScale = bco.BubbleScale;
            ShowNegativeBubbles = bco.ShowNegativeBubbles;
            SizeRepresents = bco.SizeRepresents;
        }

        internal void MergeOptions(SLStockChartOptions sco)
        {
            HasDropLines = sco.HasDropLines;
            DropLines = sco.DropLines.Clone();
            HasHighLowLines = sco.HasHighLowLines;
            HighLowLines = sco.HighLowLines.Clone();
            HasUpDownBars = sco.HasUpDownBars;
            UpDownBars = sco.UpDownBars.Clone();
        }

        internal SLChartOptions Clone()
        {
            var co = new SLChartOptions(SeriesLinesShapeProperties.listThemeColors);
            co.BarDirection = BarDirection;
            co.BarGrouping = BarGrouping;
            co.VaryColors = VaryColors;
            co.iGapWidth = iGapWidth;
            co.iGapDepth = iGapDepth;
            co.byOverlap = byOverlap;
            co.Shape = Shape;
            co.Grouping = Grouping;
            co.ShowMarker = ShowMarker;
            co.Smooth = Smooth;
            co.iFirstSliceAngle = iFirstSliceAngle;
            co.byHoleSize = byHoleSize;
            co.HasSplit = HasSplit;
            co.SplitType = SplitType;
            co.SplitPosition = SplitPosition;

            co.SecondPiePoints = new List<int>();
            for (var i = 0; i < SecondPiePoints.Count; ++i)
                co.SecondPiePoints.Add(SecondPiePoints[i]);

            co.iSecondPieSize = iSecondPieSize;
            co.SeriesLinesShapeProperties = SeriesLinesShapeProperties.Clone();
            co.ScatterStyle = ScatterStyle;
            co.bWireframe = bWireframe;
            co.RadarStyle = RadarStyle;
            co.Bubble3D = Bubble3D;
            co.iBubbleScale = iBubbleScale;
            co.ShowNegativeBubbles = ShowNegativeBubbles;
            co.SizeRepresents = SizeRepresents;

            co.HasDropLines = HasDropLines;
            co.DropLines = DropLines.Clone();
            co.HasHighLowLines = HasHighLowLines;
            co.HighLowLines = HighLowLines.Clone();
            co.HasUpDownBars = HasUpDownBars;
            co.UpDownBars = UpDownBars.Clone();

            return co;
        }
    }
}