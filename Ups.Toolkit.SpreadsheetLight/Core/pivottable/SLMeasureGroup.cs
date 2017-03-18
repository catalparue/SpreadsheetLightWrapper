using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.pivottable
{
    internal class SLMeasureGroup
    {
        internal SLMeasureGroup()
        {
            SetAllNull();
        }

        internal string Name { get; set; }
        internal string Caption { get; set; }

        private void SetAllNull()
        {
            Name = "";
            Caption = "";
        }

        internal void FromMeasureGroup(MeasureGroup mg)
        {
            SetAllNull();

            if (mg.Name != null) Name = mg.Name.Value;
            if (mg.Caption != null) Caption = mg.Caption.Value;
        }

        internal MeasureGroup ToMeasureGroup()
        {
            var mg = new MeasureGroup();
            mg.Name = Name;
            mg.Caption = Caption;

            return mg;
        }

        internal SLMeasureGroup Clone()
        {
            var mg = new SLMeasureGroup();
            mg.Name = Name;
            mg.Caption = Caption;

            return mg;
        }
    }
}