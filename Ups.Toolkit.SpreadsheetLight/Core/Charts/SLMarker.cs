using System.Collections.Generic;
using System.Drawing;
using Ups.Toolkit.SpreadsheetLight.Core.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Ups.Toolkit.SpreadsheetLight.Core.Charts
{
    /// <summary>
    ///     Encapsulates properties and methods for setting data markers in charts.
    ///     This simulates the DocumentFormat.OpenXml.Drawing.Charts.Marker class.
    /// </summary>
    public class SLMarker
    {
        internal byte? bySize;

        internal C.MarkerStyleValues? vSymbol;

        internal SLMarker(List<Color> ThemeColors)
        {
            vSymbol = null;
            bySize = null;
            ShapeProperties = new SLShapeProperties(ThemeColors);
        }

        internal bool HasMarker
        {
            get { return (vSymbol != null) || (bySize != null) || ShapeProperties.HasShapeProperties; }
        }

        /// <summary>
        ///     Marker symbol.
        /// </summary>
        public C.MarkerStyleValues Symbol
        {
            get { return vSymbol ?? C.MarkerStyleValues.Auto; }
            set { vSymbol = value; }
        }

        /// <summary>
        ///     Range is 2 to 72 inclusive. Default is 5 in Open XML but Excel uses 7.
        /// </summary>
        public byte Size
        {
            get { return bySize ?? 5; }
            set
            {
                bySize = value;
                if (bySize != null)
                {
                    if (bySize < 2) bySize = 2;
                    if (bySize > 72) bySize = 72;
                }
            }
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
        ///     Line properties.
        /// </summary>
        public SLLinePropertiesType Line
        {
            get { return ShapeProperties.Outline; }
        }

        internal C.Marker ToMarker(bool IsStylish = false)
        {
            var m = new C.Marker();
            if (vSymbol != null) m.Symbol = new C.Symbol {Val = vSymbol.Value};
            if (bySize != null) m.Size = new C.Size {Val = bySize.Value};

            if (ShapeProperties.HasShapeProperties)
                m.ChartShapeProperties = ShapeProperties.ToChartShapeProperties(IsStylish);

            return m;
        }

        internal SLMarker Clone()
        {
            var m = new SLMarker(ShapeProperties.listThemeColors);
            m.Symbol = Symbol;
            m.bySize = bySize;
            m.ShapeProperties = ShapeProperties.Clone();

            return m;
        }
    }
}