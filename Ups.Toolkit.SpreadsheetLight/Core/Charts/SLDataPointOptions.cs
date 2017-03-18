using System.Collections.Generic;
using System.Drawing;
using Ups.Toolkit.SpreadsheetLight.Core.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Ups.Toolkit.SpreadsheetLight.Core.Charts
{
    /// <summary>
    ///     Encapsulates properties and methods for setting data point options for charts.
    /// </summary>
    public class SLDataPointOptions
    {
        internal bool? bBubble3D;

        // "default" is 25%, range of 0% to 400%
        // but we're not enforcing the range
        internal uint? iExplosion;

        // pictureoptions?

        internal SLDataPointOptions(List<Color> ThemeColors)
        {
            ShapeProperties = new SLShapeProperties(ThemeColors);
            InvertIfNegative = null;
            Marker = new SLMarker(ThemeColors);
            iExplosion = null;
            bBubble3D = null;
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

        internal C.DataPoint ToDataPoint(int index, bool IsStylish = false)
        {
            var pt = new C.DataPoint();

            pt.Index = new C.Index {Val = (uint) index};

            if (Marker.HasMarker) pt.Marker = Marker.ToMarker(IsStylish);

            if (bBubble3D != null) pt.Bubble3D = new C.Bubble3D {Val = bBubble3D.Value};

            if (iExplosion != null) pt.Explosion = new C.Explosion {Val = iExplosion.Value};

            if (ShapeProperties.HasShapeProperties)
                pt.ChartShapeProperties = ShapeProperties.ToChartShapeProperties(IsStylish);

            return pt;
        }

        internal SLDataPointOptions Clone()
        {
            var dpo = new SLDataPointOptions(ShapeProperties.listThemeColors);
            dpo.ShapeProperties = ShapeProperties.Clone();
            dpo.InvertIfNegative = InvertIfNegative;
            dpo.Marker = Marker.Clone();
            dpo.iExplosion = iExplosion;
            dpo.bBubble3D = bBubble3D;

            return dpo;
        }
    }
}