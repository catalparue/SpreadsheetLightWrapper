using System;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLightWrapper.Core.misc;

namespace SpreadsheetLightWrapper.Core.style
{
    /// <summary>
    ///     Encapsulates properties and methods for specifying cell styles. This simulates the
    ///     DocumentFormat.OpenXml.Spreadsheet.CellStyle class.
    /// </summary>
    public class SLCellStyle
    {
        /// <summary>
        ///     Initializes an instance of SLCellStyle.
        /// </summary>
        public SLCellStyle()
        {
            SetAllNull();
        }

        /// <summary>
        ///     SheetName of the cell style.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        ///     Specifies a zero-based index referencing a CellFormat in the CellStyleFormats class.
        /// </summary>
        public uint FormatId { get; set; }

        /// <summary>
        ///     Specifies the index of a built-in cell style.
        /// </summary>
        public uint? BuiltinId { get; set; }

        /// <summary>
        ///     Specifies that the formatting is for an outline style.
        /// </summary>
        public uint? OutlineLevel { get; set; }

        /// <summary>
        ///     Specifies if the style is shown in the application user interface.
        /// </summary>
        public bool? Hidden { get; set; }

        /// <summary>
        ///     Specifies if the built-in cell style is customized.
        /// </summary>
        public bool? CustomBuiltin { get; set; }

        private void SetAllNull()
        {
            Name = null;
            FormatId = 0;
            BuiltinId = null;
            OutlineLevel = null;
            Hidden = null;
            CustomBuiltin = null;
        }

        internal void FromCellStyle(CellStyle cs)
        {
            SetAllNull();

            if (cs.Name != null) Name = cs.Name.Value;
            if (cs.FormatId != null) FormatId = cs.FormatId.Value;
            if (cs.BuiltinId != null) BuiltinId = cs.BuiltinId.Value;
            if (cs.OutlineLevel != null) OutlineLevel = cs.OutlineLevel.Value;
            if (cs.Hidden != null) Hidden = cs.Hidden.Value;
            if (cs.CustomBuiltin != null) CustomBuiltin = cs.CustomBuiltin.Value;
        }

        internal CellStyle ToCellStyle()
        {
            var cs = new CellStyle();
            if (Name != null) cs.Name = Name;
            cs.FormatId = FormatId;
            if (BuiltinId != null) cs.BuiltinId = BuiltinId.Value;
            if (OutlineLevel != null) cs.OutlineLevel = OutlineLevel.Value;
            if (Hidden != null) cs.Hidden = Hidden.Value;
            if (CustomBuiltin != null) cs.CustomBuiltin = CustomBuiltin.Value;

            return cs;
        }

        internal void FromHash(string Hash)
        {
            SetAllNull();
            var sa = Hash.Split(new[] {SLConstants.XmlCellStyleAttributeSeparator}, StringSplitOptions.None);

            if (sa.Length >= 6)
            {
                // weird if the actual name *is* "null"...
                if (!sa[0].Equals("null")) Name = sa[0];
                else Name = string.Empty;

                FormatId = uint.Parse(sa[1]);

                if (!sa[2].Equals("null")) BuiltinId = uint.Parse(sa[2]);

                if (!sa[3].Equals("null")) OutlineLevel = uint.Parse(sa[3]);

                if (!sa[4].Equals("null"))
                    if (sa[4].Equals("true")) Hidden = true;
                    else if (sa[4].Equals("false")) Hidden = false;

                if (!sa[5].Equals("null"))
                    if (sa[5].Equals("true")) CustomBuiltin = true;
                    else if (sa[5].Equals("false")) CustomBuiltin = false;
            }
        }

        internal string ToHash()
        {
            var sb = new StringBuilder();

            if (Name != null) sb.AppendFormat("{0}{1}", Name, SLConstants.XmlCellStyleAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlCellStyleAttributeSeparator);

            sb.AppendFormat("{0}{1}", FormatId, SLConstants.XmlCellStyleAttributeSeparator);

            if (BuiltinId != null)
                sb.AppendFormat("{0}{1}", BuiltinId.Value, SLConstants.XmlCellStyleAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlCellStyleAttributeSeparator);

            if (OutlineLevel != null)
                sb.AppendFormat("{0}{1}", OutlineLevel.Value, SLConstants.XmlCellStyleAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlCellStyleAttributeSeparator);

            if (Hidden != null)
                sb.AppendFormat("{0}{1}", Hidden.Value ? "true" : "false", SLConstants.XmlCellStyleAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlCellStyleAttributeSeparator);

            if (CustomBuiltin != null)
                sb.AppendFormat("{0}{1}", CustomBuiltin.Value ? "true" : "false",
                    SLConstants.XmlCellStyleAttributeSeparator);
            else sb.AppendFormat("null{0}", SLConstants.XmlCellStyleAttributeSeparator);

            return sb.ToString();
        }

        internal string WriteToXmlTag()
        {
            var sb = new StringBuilder();
            sb.Append("<x:cellStyle");
            if (Name != null) sb.AppendFormat(" name=\"{0}\"", Name);
            sb.AppendFormat(" xfId=\"{0}\"", FormatId);
            if (BuiltinId != null) sb.AppendFormat(" builtinId=\"{0}\"", BuiltinId.Value);
            if (OutlineLevel != null) sb.AppendFormat(" iLevel=\"{0}\"", OutlineLevel.Value);
            if (Hidden != null) sb.AppendFormat(" hidden=\"{0}\"", Hidden.Value);
            if (CustomBuiltin != null) sb.AppendFormat(" customBuiltin=\"{0}\"", CustomBuiltin.Value);
            sb.Append(" />");

            return sb.ToString();
        }

        internal SLCellStyle Clone()
        {
            var cs = new SLCellStyle();
            cs.Name = Name;
            cs.FormatId = FormatId;
            cs.BuiltinId = BuiltinId;
            cs.OutlineLevel = OutlineLevel;
            cs.Hidden = Hidden;
            cs.CustomBuiltin = CustomBuiltin;

            return cs;
        }
    }
}