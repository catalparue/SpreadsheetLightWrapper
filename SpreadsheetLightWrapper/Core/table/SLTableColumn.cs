using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.table
{
    internal class SLTableColumn
    {
        internal bool HasCalculatedColumnFormula;

        internal bool HasTotalsRowFormula;

        internal bool HasTotalsRowFunction;

        internal bool HasXmlColumnProperties;
        private TotalsRowFunctionValues vTotalsRowFunction;

        internal SLTableColumn()
        {
            SetAllNull();
        }

        internal SLCalculatedColumnFormula CalculatedColumnFormula { get; set; }
        internal SLTotalsRowFormula TotalsRowFormula { get; set; }
        internal SLXmlColumnProperties XmlColumnProperties { get; set; }

        internal uint Id { get; set; }
        internal string UniqueName { get; set; }
        internal string Name { get; set; }

        internal TotalsRowFunctionValues TotalsRowFunction
        {
            get { return vTotalsRowFunction; }
            set
            {
                vTotalsRowFunction = value;
                HasTotalsRowFunction = vTotalsRowFunction != TotalsRowFunctionValues.None ? true : false;
            }
        }

        internal string TotalsRowLabel { get; set; }
        internal uint? QueryTableFieldId { get; set; }
        internal uint? HeaderRowDifferentialFormattingId { get; set; }
        internal uint? DataFormatId { get; set; }
        internal uint? TotalsRowDifferentialFormattingId { get; set; }
        internal string HeaderRowCellStyle { get; set; }
        internal string DataCellStyle { get; set; }
        internal string TotalsRowCellStyle { get; set; }

        private void SetAllNull()
        {
            CalculatedColumnFormula = new SLCalculatedColumnFormula();
            HasCalculatedColumnFormula = false;
            TotalsRowFormula = new SLTotalsRowFormula();
            HasTotalsRowFormula = false;
            XmlColumnProperties = new SLXmlColumnProperties();
            HasXmlColumnProperties = false;

            Id = 0;
            UniqueName = null;
            Name = string.Empty;
            TotalsRowFunction = TotalsRowFunctionValues.None;
            HasTotalsRowFunction = false;
            TotalsRowLabel = null;
            QueryTableFieldId = null;
            HeaderRowDifferentialFormattingId = null;
            DataFormatId = null;
            TotalsRowDifferentialFormattingId = null;
            HeaderRowCellStyle = null;
            DataCellStyle = null;
            TotalsRowCellStyle = null;
        }

        internal void FromTableColumn(TableColumn tc)
        {
            SetAllNull();

            if (tc.CalculatedColumnFormula != null)
            {
                HasCalculatedColumnFormula = true;
                CalculatedColumnFormula.FromCalculatedColumnFormula(tc.CalculatedColumnFormula);
            }
            if (tc.TotalsRowFormula != null)
            {
                HasTotalsRowFormula = true;
                TotalsRowFormula.FromTotalsRowFormula(tc.TotalsRowFormula);
            }
            if (tc.XmlColumnProperties != null)
            {
                HasXmlColumnProperties = true;
                XmlColumnProperties.FromXmlColumnProperties(tc.XmlColumnProperties);
            }

            Id = tc.Id.Value;
            if (tc.UniqueName != null) UniqueName = tc.UniqueName.Value;
            Name = tc.Name.Value;

            if (tc.TotalsRowFunction != null) TotalsRowFunction = tc.TotalsRowFunction.Value;
            if (tc.TotalsRowLabel != null) TotalsRowLabel = tc.TotalsRowLabel.Value;
            if (tc.QueryTableFieldId != null) QueryTableFieldId = tc.QueryTableFieldId.Value;
            if (tc.HeaderRowDifferentialFormattingId != null)
                HeaderRowDifferentialFormattingId = tc.HeaderRowDifferentialFormattingId.Value;
            if (tc.DataFormatId != null) DataFormatId = tc.DataFormatId.Value;
            if (tc.TotalsRowDifferentialFormattingId != null)
                TotalsRowDifferentialFormattingId = tc.TotalsRowDifferentialFormattingId.Value;
            if (tc.HeaderRowCellStyle != null) HeaderRowCellStyle = tc.HeaderRowCellStyle.Value;
            if (tc.DataCellStyle != null) DataCellStyle = tc.DataCellStyle.Value;
            if (tc.TotalsRowCellStyle != null) TotalsRowCellStyle = tc.TotalsRowCellStyle.Value;
        }

        internal TableColumn ToTableColumn()
        {
            var tc = new TableColumn();
            if (HasCalculatedColumnFormula)
                tc.CalculatedColumnFormula = CalculatedColumnFormula.ToCalculatedColumnFormula();
            if (HasTotalsRowFormula)
                tc.TotalsRowFormula = TotalsRowFormula.ToTotalsRowFormula();
            if (HasXmlColumnProperties)
                tc.XmlColumnProperties = XmlColumnProperties.ToXmlColumnProperties();

            tc.Id = Id;
            if (UniqueName != null) tc.UniqueName = UniqueName;
            tc.Name = Name;

            if (HasTotalsRowFunction) tc.TotalsRowFunction = TotalsRowFunction;
            if (TotalsRowLabel != null) tc.TotalsRowLabel = TotalsRowLabel;
            if (QueryTableFieldId != null) tc.QueryTableFieldId = QueryTableFieldId.Value;
            if (HeaderRowDifferentialFormattingId != null)
                tc.HeaderRowDifferentialFormattingId = HeaderRowDifferentialFormattingId.Value;
            if (DataFormatId != null) tc.DataFormatId = DataFormatId.Value;
            if (TotalsRowDifferentialFormattingId != null)
                tc.TotalsRowDifferentialFormattingId = TotalsRowDifferentialFormattingId.Value;
            if (HeaderRowCellStyle != null) tc.HeaderRowCellStyle = HeaderRowCellStyle;
            if (DataCellStyle != null) tc.DataCellStyle = DataCellStyle;
            if (TotalsRowCellStyle != null) tc.TotalsRowCellStyle = TotalsRowCellStyle;

            return tc;
        }

        internal SLTableColumn Clone()
        {
            var tc = new SLTableColumn();
            tc.HasCalculatedColumnFormula = HasCalculatedColumnFormula;
            tc.CalculatedColumnFormula = CalculatedColumnFormula.Clone();
            tc.HasTotalsRowFormula = HasTotalsRowFormula;
            tc.TotalsRowFormula = TotalsRowFormula.Clone();
            tc.HasXmlColumnProperties = HasXmlColumnProperties;
            tc.XmlColumnProperties = XmlColumnProperties.Clone();
            tc.Id = Id;
            tc.UniqueName = UniqueName;
            tc.Name = Name;
            tc.HasTotalsRowFunction = HasTotalsRowFunction;
            tc.vTotalsRowFunction = vTotalsRowFunction;
            tc.TotalsRowLabel = TotalsRowLabel;
            tc.QueryTableFieldId = QueryTableFieldId;
            tc.HeaderRowDifferentialFormattingId = HeaderRowDifferentialFormattingId;
            tc.DataFormatId = DataFormatId;
            tc.TotalsRowDifferentialFormattingId = TotalsRowDifferentialFormattingId;
            tc.HeaderRowCellStyle = HeaderRowCellStyle;
            tc.DataCellStyle = DataCellStyle;
            tc.TotalsRowCellStyle = TotalsRowCellStyle;

            return tc;
        }
    }
}