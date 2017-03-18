using System;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;
using Ups.Toolkit.SpreadsheetLight.Core.misc;

namespace Ups.Toolkit.SpreadsheetLight.Core.style
{
    internal class SLTableStyle
    {
        internal string Name;
        internal string TableStyleInnerXml;

        internal SLTableStyle()
        {
            SetAllNull();
        }

        internal bool? Pivot { get; set; }

        internal bool? Table { get; set; }

        internal uint? Count { get; set; }

        private void SetAllNull()
        {
            TableStyleInnerXml = string.Empty;
            Name = string.Empty;
            Pivot = null;
            Table = null;
            Count = null;
        }

        internal void FromTableStyle(TableStyle ts)
        {
            SetAllNull();

            TableStyleInnerXml = ts.InnerXml;

            // this is a required field, so it can't be null, but just in case...
            if (ts.Name != null) Name = ts.Name.Value;
            else Name = string.Empty;

            if (ts.Pivot != null)
                Pivot = ts.Pivot.Value;

            if (ts.Table != null)
                Table = ts.Table.Value;

            if (ts.Count != null)
                Count = ts.Count.Value;
        }

        internal TableStyle ToTableStyle()
        {
            var ts = new TableStyle();
            ts.InnerXml = SLTool.RemoveNamespaceDeclaration(TableStyleInnerXml);
            ts.Name = Name;

            if (Pivot != null) ts.Pivot = Pivot.Value;
            if (Table != null) ts.Table = Table.Value;
            if (Count != null) ts.Count = Count.Value;

            return ts;
        }

        internal void FromHash(string Hash)
        {
            var ts = new TableStyle();

            var saElementAttribute = Hash.Split(new[] {SLConstants.XmlTableStyleElementAttributeSeparator},
                StringSplitOptions.None);

            if (saElementAttribute.Length >= 2)
            {
                ts.InnerXml = saElementAttribute[0];
                var sa = saElementAttribute[1].Split(new[] {SLConstants.XmlTableStyleAttributeSeparator},
                    StringSplitOptions.None);
                if (sa.Length >= 4)
                {
                    ts.Name = sa[0];

                    if (!sa[1].Equals("null")) ts.Pivot = bool.Parse(sa[1]);

                    if (!sa[2].Equals("null")) ts.Table = bool.Parse(sa[2]);

                    if (!sa[3].Equals("null")) ts.Count = uint.Parse(sa[3]);
                }
            }

            FromTableStyle(ts);
        }

        internal string ToHash()
        {
            var ts = ToTableStyle();
            var sXml = SLTool.RemoveNamespaceDeclaration(ts.InnerXml);

            var sb = new StringBuilder();

            sb.AppendFormat("{0}{1}", sXml, SLConstants.XmlTableStyleElementAttributeSeparator);

            sb.AppendFormat("{0}{1}", ts.Name.Value, SLConstants.XmlTableStyleAttributeSeparator);

            if (ts.Pivot != null)
                sb.AppendFormat("{0}{1}", ts.Pivot.Value, SLConstants.XmlTableStyleAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlTableStyleAttributeSeparator);

            if (ts.Table != null)
                sb.AppendFormat("{0}{1}", ts.Table.Value, SLConstants.XmlTableStyleAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlTableStyleAttributeSeparator);

            if (ts.Count != null)
                sb.AppendFormat("{0}{1}", ts.Count.Value, SLConstants.XmlTableStyleAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlTableStyleAttributeSeparator);

            return sb.ToString();
        }

        internal string WriteToXmlTag()
        {
            var sb = new StringBuilder();
            sb.AppendFormat("<x:tableStyle name=\"{0}\"", Name);
            if ((Pivot != null) && !Pivot.Value) sb.Append(" pivot=\"0\"");
            if ((Table != null) && !Table.Value) sb.Append(" table=\"0\"");
            if (Count != null) sb.AppendFormat(" count=\"{0}\"", Count.Value);

            if (TableStyleInnerXml.Length > 0)
            {
                sb.Append(">");
                sb.Append(TableStyleInnerXml);
                sb.Append("</x:tableStyle>");
            }
            else
            {
                sb.Append(" />");
            }

            return sb.ToString();
        }

        internal SLTableStyle Clone()
        {
            var ts = new SLTableStyle();
            ts.TableStyleInnerXml = TableStyleInnerXml;
            ts.Name = Name;
            ts.Pivot = Pivot;
            ts.Table = Table;
            ts.Count = Count;

            return ts;
        }
    }
}