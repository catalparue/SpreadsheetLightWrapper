using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.pivottable
{
    internal class SLServerFormat
    {
        internal SLServerFormat()
        {
            SetAllNull();
        }

        internal string Culture { get; set; }
        internal string Format { get; set; }

        private void SetAllNull()
        {
            Culture = "";
            Format = "";
        }

        internal void FromServerFormat(ServerFormat sf)
        {
            SetAllNull();

            if (sf.Culture != null) Culture = sf.Culture.Value;
            if (sf.Format != null) Format = sf.Format.Value;
        }

        internal ServerFormat ToServerFormat()
        {
            var sf = new ServerFormat();
            if ((Culture != null) && (Culture.Length > 0)) sf.Culture = Culture;
            if ((Format != null) && (Format.Length > 0)) sf.Format = Format;

            return sf;
        }

        internal SLServerFormat Clone()
        {
            var sf = new SLServerFormat();
            sf.Culture = Culture;
            sf.Format = Format;

            return sf;
        }
    }
}