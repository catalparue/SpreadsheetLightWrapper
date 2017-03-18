using System;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;
using Ups.Toolkit.SpreadsheetLight.Core.misc;

namespace Ups.Toolkit.SpreadsheetLight.Core.style
{
    /// <summary>
    ///     Encapsulates properties and methods for cell content protection. The properties don't take effect unless the
    ///     worksheet is protected. This simulates the DocumentFormat.OpenXml.Spreadsheet.Protection class.
    /// </summary>
    public class SLProtection
    {
        /// <summary>
        ///     Initializes an instance of SLProtection.
        /// </summary>
        public SLProtection()
        {
            SetAllNull();
        }

        /// <summary>
        ///     Specifies if the cell is locked. If locked and the worksheet is protected, then the worksheet's protection options
        ///     are ignored.
        /// </summary>
        public bool? Locked { get; set; }

        /// <summary>
        ///     Specifies if the cell is hidden. If hidden and the worksheet is protected, then cell contents are hidden and only
        ///     cell values are displayed. For example, the cell formula is hidden, but the value of the cell formula is still
        ///     displayed.
        /// </summary>
        public bool? Hidden { get; set; }

        private void SetAllNull()
        {
            Locked = null;
            Hidden = null;
        }

        internal void FromProtection(Protection p)
        {
            SetAllNull();

            if (p.Locked != null)
                Locked = p.Locked.Value;

            if (p.Hidden != null)
                Hidden = p.Hidden.Value;
        }

        internal Protection ToProtection()
        {
            var p = new Protection();
            if (Locked != null) p.Locked = Locked.Value;
            if (Hidden != null) p.Hidden = Hidden.Value;

            return p;
        }

        internal void FromHash(string Hash)
        {
            SetAllNull();

            var sa = Hash.Split(new[] {SLConstants.XmlProtectionAttributeSeparator}, StringSplitOptions.None);
            if (sa.Length >= 2)
            {
                if (!sa[0].Equals("null")) Locked = bool.Parse(sa[0]);

                if (!sa[1].Equals("null")) Hidden = bool.Parse(sa[1]);
            }
        }

        internal string ToHash()
        {
            var sb = new StringBuilder();

            if (Locked != null) sb.AppendFormat("{0}{1}", Locked.Value, SLConstants.XmlProtectionAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlProtectionAttributeSeparator);

            if (Hidden != null) sb.AppendFormat("{0}{1}", Hidden.Value, SLConstants.XmlProtectionAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlProtectionAttributeSeparator);

            return sb.ToString();
        }

        internal string WriteToXmlTag()
        {
            var sb = new StringBuilder();
            sb.Append("<x:protection");
            if (Locked != null) sb.AppendFormat(" locked=\"{0}\"", Locked.Value ? "1" : "0");
            if (Hidden != null) sb.AppendFormat(" hidden=\"{0}\"", Hidden.Value ? "1" : "0");
            sb.Append(" />");

            return sb.ToString();
        }

        internal SLProtection Clone()
        {
            var p = new SLProtection();
            p.Locked = Locked;
            p.Hidden = Hidden;

            return p;
        }
    }
}