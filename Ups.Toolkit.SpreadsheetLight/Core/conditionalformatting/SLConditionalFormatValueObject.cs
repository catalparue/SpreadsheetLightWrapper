using DocumentFormat.OpenXml.Spreadsheet;
using Ups.Toolkit.SpreadsheetLight.Core.office2010;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace Ups.Toolkit.SpreadsheetLight.Core.conditionalformatting
{
    internal class SLConditionalFormatValueObject
    {
        internal SLConditionalFormatValueObject()
        {
            SetAllNull();
        }

        internal ConditionalFormatValueObjectValues Type { get; set; }
        internal string Val { get; set; }
        internal bool GreaterThanOrEqual { get; set; }

        private void SetAllNull()
        {
            Type = ConditionalFormatValueObjectValues.Percentile;
            Val = string.Empty;
            GreaterThanOrEqual = true;
        }

        internal void FromConditionalFormatValueObject(ConditionalFormatValueObject cfvo)
        {
            SetAllNull();

            Type = cfvo.Type.Value;
            if (cfvo.Val != null) Val = cfvo.Val.Value;
            if (cfvo.GreaterThanOrEqual != null) GreaterThanOrEqual = cfvo.GreaterThanOrEqual.Value;
        }

        internal ConditionalFormatValueObject ToConditionalFormatValueObject()
        {
            var cfvo = new ConditionalFormatValueObject();
            cfvo.Type = Type;

            if (Val.Length > 0)
                if (Val.StartsWith("=")) cfvo.Val = Val.Substring(1);
                else cfvo.Val = Val;

            if (!GreaterThanOrEqual) cfvo.GreaterThanOrEqual = false;

            return cfvo;
        }

        internal SLConditionalFormattingValueObject2010 ToSLConditionalFormattingValueObject2010()
        {
            var cfvo2010 = new SLConditionalFormattingValueObject2010();
            cfvo2010.Formula = Val;

            switch (Type)
            {
                case ConditionalFormatValueObjectValues.Formula:
                    cfvo2010.Type = X14.ConditionalFormattingValueObjectTypeValues.Formula;
                    break;
                case ConditionalFormatValueObjectValues.Max:
                    cfvo2010.Type = X14.ConditionalFormattingValueObjectTypeValues.Max;
                    break;
                case ConditionalFormatValueObjectValues.Min:
                    cfvo2010.Type = X14.ConditionalFormattingValueObjectTypeValues.Min;
                    break;
                case ConditionalFormatValueObjectValues.Number:
                    cfvo2010.Type = X14.ConditionalFormattingValueObjectTypeValues.Numeric;
                    break;
                case ConditionalFormatValueObjectValues.Percent:
                    cfvo2010.Type = X14.ConditionalFormattingValueObjectTypeValues.Percent;
                    break;
                case ConditionalFormatValueObjectValues.Percentile:
                    cfvo2010.Type = X14.ConditionalFormattingValueObjectTypeValues.Percentile;
                    break;
            }

            cfvo2010.GreaterThanOrEqual = GreaterThanOrEqual;

            return cfvo2010;
        }

        internal SLConditionalFormatValueObject Clone()
        {
            var cfvo = new SLConditionalFormatValueObject();
            cfvo.Type = Type;
            cfvo.Val = Val;
            cfvo.GreaterThanOrEqual = GreaterThanOrEqual;

            return cfvo;
        }
    }
}