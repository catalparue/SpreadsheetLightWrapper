using System.Collections.Generic;
using System.Drawing;
using Ups.Toolkit.SpreadsheetLight.Core.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Ups.Toolkit.SpreadsheetLight.Core.Charts
{
    /// <summary>
    ///     Data series customization options. Note that all supported chart data series properties are available, but only the
    ///     relevant properties (to chart type) will be used.
    /// </summary>
    public class SLDataSeriesOptions
    {
        internal bool? bBubble3D;

        // "default" is 25%, range of 0% to 400%
        // but we're not enforcing the range
        internal uint? iExplosion;

        internal C.ShapeValues? vShape;

        /// <summary>
        ///     Initializes an instance of SLDataSeriesOptions. It is recommended to use SLChart.GetDataSeriesOptions().
        /// </summary>
        public SLDataSeriesOptions()
        {
            Initialize(new List<Color>());
        }

        internal SLDataSeriesOptions(List<Color> ThemeColors)
        {
            Initialize(ThemeColors);
        }

        internal SLShapeProperties ShapeProperties { get; set; }

        /// <summary>
        ///     Fill properties.
        /// </summary>
        public SLFill Fill
        {
            get { return ShapeProperties.Fill; }
        }

        /// <summary>
        ///     Border/Line properties.
        /// </summary>
        public SLLinePropertiesType Line
        {
            get { return ShapeProperties.Outline; }
        }

        /// <summary>
        ///     Shadow properties.
        /// </summary>
        public SLShadowEffect Shadow
        {
            get { return ShapeProperties.EffectList.Shadow; }
        }

        /// <summary>
        ///     Glow properties.
        /// </summary>
        public SLGlow Glow
        {
            get { return ShapeProperties.EffectList.Glow; }
        }

        /// <summary>
        ///     Soft edge properties.
        /// </summary>
        public SLSoftEdge SoftEdge
        {
            get { return ShapeProperties.EffectList.SoftEdge; }
        }

        /// <summary>
        ///     3D format properties.
        /// </summary>
        public SLFormat3D Format3D
        {
            get { return ShapeProperties.Format3D; }
        }

        // internally, the default is actually true in Open XML, but when null it's false.
        // The Open XML docs state it's supposed to be true when the tag is missing. I don't know...
        /// <summary>
        ///     Invert colors if negative. If null, the effective default is used (false). This is for bar charts, column charts
        ///     and bubble charts.
        /// </summary>
        public bool? InvertIfNegative { get; set; }

        /// <summary>
        ///     Marker properties. This is for line charts, radar charts and scatter charts.
        /// </summary>
        public SLMarker Marker { get; set; }

        /// <summary>
        ///     The explosion distance from the center of the pie in percentage. It is suggested to keep the range between 0% and
        ///     400%.
        /// </summary>
        public uint Explosion
        {
            get { return iExplosion ?? 0; }
            set { iExplosion = value; }
        }

        internal bool Bubble3D
        {
            get { return bBubble3D ?? true; }
            set { bBubble3D = value; }
        }

        /// <summary>
        ///     Whether the line connecting data points use C splines (instead of straight lines). This is for line charts and
        ///     scatter charts.
        /// </summary>
        public bool Smooth { get; set; }

        /// <summary>
        ///     The shape of data series for 3D bar and column charts.
        /// </summary>
        public C.ShapeValues Shape
        {
            get { return vShape ?? C.ShapeValues.Box; }
            set { vShape = value; }
        }

        private void Initialize(List<Color> ThemeColors)
        {
            ShapeProperties = new SLShapeProperties(ThemeColors);
            InvertIfNegative = null;
            Marker = new SLMarker(ThemeColors);
            iExplosion = null;
            bBubble3D = null;
            Smooth = false;
            vShape = null;
        }

        internal SLDataSeriesOptions Clone()
        {
            var dso = new SLDataSeriesOptions(ShapeProperties.listThemeColors);
            dso.ShapeProperties = ShapeProperties.Clone();
            dso.InvertIfNegative = InvertIfNegative;
            dso.Marker = Marker.Clone();
            dso.iExplosion = iExplosion;
            dso.bBubble3D = bBubble3D;
            dso.Smooth = Smooth;
            dso.vShape = vShape;

            return dso;
        }
    }
}