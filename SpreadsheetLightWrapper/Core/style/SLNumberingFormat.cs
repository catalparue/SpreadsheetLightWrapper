using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.style
{
    internal class SLNumberingFormat
    {
        internal SLNumberingFormat()
        {
            SetAllNull();
        }

        internal uint NumberFormatId { get; set; }

        internal string FormatCode { get; set; }

        private void SetAllNull()
        {
            NumberFormatId = 0;
            FormatCode = string.Empty;
        }

        internal void FromNumberingFormat(NumberingFormat nf)
        {
            SetAllNull();

            if (nf.NumberFormatId != null)
                NumberFormatId = nf.NumberFormatId.Value;
            else
                NumberFormatId = 0;

            if (nf.FormatCode != null)
                FormatCode = nf.FormatCode.Value;
            else
                FormatCode = string.Empty;
        }

        internal NumberingFormat ToNumberingFormat()
        {
            var nf = new NumberingFormat();
            nf.NumberFormatId = NumberFormatId;
            nf.FormatCode = FormatCode;

            return nf;
        }

        internal void FromHash(string Hash)
        {
            FormatCode = Hash;
        }

        internal string ToHash()
        {
            return FormatCode;
        }

        internal SLNumberingFormat Clone()
        {
            var nf = new SLNumberingFormat();
            nf.NumberFormatId = NumberFormatId;
            nf.FormatCode = FormatCode;

            return nf;
        }
    }
}