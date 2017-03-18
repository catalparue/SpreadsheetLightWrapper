using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.table
{
    internal class SLXmlColumnProperties
    {
        internal SLXmlColumnProperties()
        {
            SetAllNull();
        }

        internal uint MapId { get; set; }
        internal string XPath { get; set; }
        internal bool? Denormalized { get; set; }
        internal XmlDataValues XmlDataType { get; set; }

        private void SetAllNull()
        {
            MapId = 0;
            XPath = string.Empty;
            Denormalized = null;
            XmlDataType = XmlDataValues.AnyType;
        }

        internal void FromXmlColumnProperties(XmlColumnProperties xcp)
        {
            SetAllNull();

            if (xcp.MapId != null) MapId = xcp.MapId.Value;
            if (xcp.XPath != null) XPath = xcp.XPath.Value;
            if ((xcp.Denormalized != null) && xcp.Denormalized.Value) Denormalized = xcp.Denormalized.Value;
            if (xcp.XmlDataType != null) XmlDataType = xcp.XmlDataType.Value;
        }

        internal XmlColumnProperties ToXmlColumnProperties()
        {
            var xcp = new XmlColumnProperties();
            xcp.MapId = MapId;
            xcp.XPath = XPath;
            if ((Denormalized != null) && Denormalized.Value) xcp.Denormalized = Denormalized.Value;
            xcp.XmlDataType = XmlDataType;

            return xcp;
        }

        internal SLXmlColumnProperties Clone()
        {
            var xcp = new SLXmlColumnProperties();
            xcp.MapId = MapId;
            xcp.XPath = XPath;
            xcp.Denormalized = Denormalized;
            xcp.XmlDataType = XmlDataType;

            return xcp;
        }
    }
}