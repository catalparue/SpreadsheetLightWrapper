using DocumentFormat.OpenXml.Office.Excel;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace Ups.Toolkit.SpreadsheetLight.Core.office2010
{
    internal class SLConditionalFormattingValueObject2010
    {
        internal SLConditionalFormattingValueObject2010()
        {
            SetAllNull();
        }

        //http://msdn.microsoft.com/en-us/library/documentformat.openxml.office2010.excel.conditionalformattingvalueobject.aspx

        internal string Formula { get; set; }
        internal X14.ConditionalFormattingValueObjectTypeValues Type { get; set; }
        internal bool GreaterThanOrEqual { get; set; }

        private void SetAllNull()
        {
            Formula = string.Empty;
            Type = X14.ConditionalFormattingValueObjectTypeValues.Percentile;
            GreaterThanOrEqual = true;
        }

        internal void FromConditionalFormattingValueObject(X14.ConditionalFormattingValueObject cfvo)
        {
            SetAllNull();

            if (cfvo.Formula != null) Formula = cfvo.Formula.Text;
            Type = cfvo.Type.Value;
            if (cfvo.GreaterThanOrEqual != null) GreaterThanOrEqual = cfvo.GreaterThanOrEqual.Value;
        }

        internal X14.ConditionalFormattingValueObject ToConditionalFormattingValueObject()
        {
            var cfvo = new X14.ConditionalFormattingValueObject();

            if (Formula.Length > 0)
                if (Formula.StartsWith("="))
                    cfvo.Formula = new Formula(Formula.Substring(1));
                else
                    cfvo.Formula = new Formula(Formula);
            cfvo.Type = Type;
            if (!GreaterThanOrEqual) cfvo.GreaterThanOrEqual = GreaterThanOrEqual;

            return cfvo;
        }

        internal SLConditionalFormattingValueObject2010 Clone()
        {
            var cfvo = new SLConditionalFormattingValueObject2010();
            cfvo.Formula = Formula;
            cfvo.Type = Type;
            cfvo.GreaterThanOrEqual = GreaterThanOrEqual;

            return cfvo;
        }
    }
}