using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Ups.Toolkit.SpreadsheetLight.Core.office2010;
using Ups.Toolkit.SpreadsheetLight.Core.style;
using Color = System.Drawing.Color;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace Ups.Toolkit.SpreadsheetLight.Core.conditionalformatting
{
    internal class SLDataBar
    {
        internal bool Is2010;

        internal SLDataBar()
        {
            SetAllNull();
        }

        internal SLConditionalFormatAutoMinMaxValues MinimumType { get; set; }
        internal string MinimumValue { get; set; }
        internal SLConditionalFormatAutoMinMaxValues MaximumType { get; set; }
        internal string MaximumValue { get; set; }

        internal SLColor Color { get; set; }
        internal SLColor BorderColor { get; set; }
        internal SLColor NegativeFillColor { get; set; }
        internal SLColor NegativeBorderColor { get; set; }
        internal SLColor AxisColor { get; set; }
        internal uint MinLength { get; set; }
        internal uint MaxLength { get; set; }
        internal bool ShowValue { get; set; }
        internal bool Border { get; set; }
        internal bool Gradient { get; set; }
        internal X14.DataBarDirectionValues Direction { get; set; }
        internal bool NegativeBarColorSameAsPositive { get; set; }
        internal bool NegativeBarBorderColorSameAsPositive { get; set; }
        internal X14.DataBarAxisPositionValues AxisPosition { get; set; }

        private void SetAllNull()
        {
            Is2010 = false;
            MinimumType = SLConditionalFormatAutoMinMaxValues.Percentile;
            MinimumValue = string.Empty;
            MaximumType = SLConditionalFormatAutoMinMaxValues.Percentile;
            MaximumValue = string.Empty;
            Color = new SLColor(new List<Color>(), new List<Color>());
            BorderColor = new SLColor(new List<Color>(), new List<Color>());
            NegativeFillColor = new SLColor(new List<Color>(), new List<Color>());
            NegativeBorderColor = new SLColor(new List<Color>(), new List<Color>());
            AxisColor = new SLColor(new List<Color>(), new List<Color>());
            MinLength = 10;
            MaxLength = 90;
            ShowValue = true;
        }

        internal void FromDataBar(DataBar db)
        {
            SetAllNull();

            using (var oxr = OpenXmlReader.Create(db))
            {
                var i = 0;
                SLConditionalFormatValueObject cfvo;
                while (oxr.Read())
                    if (oxr.ElementType == typeof(ConditionalFormatValueObject))
                    {
                        if (i == 0)
                        {
                            cfvo = new SLConditionalFormatValueObject();
                            cfvo.FromConditionalFormatValueObject(
                                (ConditionalFormatValueObject) oxr.LoadCurrentElement());
                            MinimumType = TranslateToAutoMinMaxValues(cfvo.Type);
                            MinimumValue = cfvo.Val;
                            ++i;
                        }
                        else if (i == 1)
                        {
                            cfvo = new SLConditionalFormatValueObject();
                            cfvo.FromConditionalFormatValueObject(
                                (ConditionalFormatValueObject) oxr.LoadCurrentElement());
                            MaximumType = TranslateToAutoMinMaxValues(cfvo.Type);
                            MaximumValue = cfvo.Val;
                            ++i;
                        }
                    }
                    else if (oxr.ElementType == typeof(DocumentFormat.OpenXml.Spreadsheet.Color))
                    {
                        Color.FromSpreadsheetColor((DocumentFormat.OpenXml.Spreadsheet.Color) oxr.LoadCurrentElement());
                    }
            }

            if (db.MinLength != null) MinLength = db.MinLength.Value;
            if (db.MaxLength != null) MaxLength = db.MaxLength.Value;
            if (db.ShowValue != null) ShowValue = db.ShowValue.Value;
        }

        internal SLConditionalFormatAutoMinMaxValues TranslateToAutoMinMaxValues(ConditionalFormatValueObjectValues Type)
        {
            var result = SLConditionalFormatAutoMinMaxValues.Percentile;
            switch (Type)
            {
                case ConditionalFormatValueObjectValues.Formula:
                    result = SLConditionalFormatAutoMinMaxValues.Formula;
                    break;
                case ConditionalFormatValueObjectValues.Max:
                    result = SLConditionalFormatAutoMinMaxValues.Value;
                    break;
                case ConditionalFormatValueObjectValues.Min:
                    result = SLConditionalFormatAutoMinMaxValues.Value;
                    break;
                case ConditionalFormatValueObjectValues.Number:
                    result = SLConditionalFormatAutoMinMaxValues.Number;
                    break;
                case ConditionalFormatValueObjectValues.Percent:
                    result = SLConditionalFormatAutoMinMaxValues.Percent;
                    break;
                case ConditionalFormatValueObjectValues.Percentile:
                    result = SLConditionalFormatAutoMinMaxValues.Percentile;
                    break;
            }

            return result;
        }

        internal DataBar ToDataBar()
        {
            var db = new DataBar();
            if (MinLength != 10) db.MinLength = MinLength;
            if (MaxLength != 90) db.MaxLength = MaxLength;
            if (!ShowValue) db.ShowValue = ShowValue;

            SLConditionalFormatValueObject cfvo;

            cfvo = new SLConditionalFormatValueObject();
            cfvo.Type = ConditionalFormatValueObjectValues.Min;
            switch (MinimumType)
            {
                case SLConditionalFormatAutoMinMaxValues.Automatic:
                    cfvo.Type = ConditionalFormatValueObjectValues.Min;
                    cfvo.Val = string.Empty;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Formula:
                    cfvo.Type = ConditionalFormatValueObjectValues.Formula;
                    cfvo.Val = MinimumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Number:
                    cfvo.Type = ConditionalFormatValueObjectValues.Number;
                    cfvo.Val = MinimumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Percent:
                    cfvo.Type = ConditionalFormatValueObjectValues.Percent;
                    cfvo.Val = MinimumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Percentile:
                    cfvo.Type = ConditionalFormatValueObjectValues.Percentile;
                    cfvo.Val = MinimumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Value:
                    cfvo.Type = ConditionalFormatValueObjectValues.Min;
                    cfvo.Val = string.Empty;
                    break;
            }
            db.Append(cfvo.ToConditionalFormatValueObject());

            cfvo = new SLConditionalFormatValueObject();
            cfvo.Type = ConditionalFormatValueObjectValues.Max;
            switch (MaximumType)
            {
                case SLConditionalFormatAutoMinMaxValues.Automatic:
                    cfvo.Type = ConditionalFormatValueObjectValues.Max;
                    cfvo.Val = string.Empty;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Formula:
                    cfvo.Type = ConditionalFormatValueObjectValues.Formula;
                    cfvo.Val = MaximumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Number:
                    cfvo.Type = ConditionalFormatValueObjectValues.Number;
                    cfvo.Val = MaximumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Percent:
                    cfvo.Type = ConditionalFormatValueObjectValues.Percent;
                    cfvo.Val = MaximumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Percentile:
                    cfvo.Type = ConditionalFormatValueObjectValues.Percentile;
                    cfvo.Val = MaximumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Value:
                    cfvo.Type = ConditionalFormatValueObjectValues.Max;
                    cfvo.Val = string.Empty;
                    break;
            }
            db.Append(cfvo.ToConditionalFormatValueObject());

            db.Append(Color.ToSpreadsheetColor());

            return db;
        }

        internal SLDataBar2010 ToDataBar2010()
        {
            var db = new SLDataBar2010();
            switch (MinimumType)
            {
                case SLConditionalFormatAutoMinMaxValues.Automatic:
                    db.Cfvo1.Type = X14.ConditionalFormattingValueObjectTypeValues.AutoMin;
                    db.Cfvo1.Formula = string.Empty;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Formula:
                    db.Cfvo1.Type = X14.ConditionalFormattingValueObjectTypeValues.Formula;
                    db.Cfvo1.Formula = MinimumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Number:
                    db.Cfvo1.Type = X14.ConditionalFormattingValueObjectTypeValues.Numeric;
                    db.Cfvo1.Formula = MinimumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Percent:
                    db.Cfvo1.Type = X14.ConditionalFormattingValueObjectTypeValues.Percent;
                    db.Cfvo1.Formula = MinimumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Percentile:
                    db.Cfvo1.Type = X14.ConditionalFormattingValueObjectTypeValues.Percentile;
                    db.Cfvo1.Formula = MinimumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Value:
                    db.Cfvo1.Type = X14.ConditionalFormattingValueObjectTypeValues.Min;
                    db.Cfvo1.Formula = string.Empty;
                    break;
            }

            switch (MaximumType)
            {
                case SLConditionalFormatAutoMinMaxValues.Automatic:
                    db.Cfvo2.Type = X14.ConditionalFormattingValueObjectTypeValues.AutoMax;
                    db.Cfvo2.Formula = string.Empty;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Formula:
                    db.Cfvo2.Type = X14.ConditionalFormattingValueObjectTypeValues.Formula;
                    db.Cfvo2.Formula = MaximumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Number:
                    db.Cfvo2.Type = X14.ConditionalFormattingValueObjectTypeValues.Numeric;
                    db.Cfvo2.Formula = MaximumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Percent:
                    db.Cfvo2.Type = X14.ConditionalFormattingValueObjectTypeValues.Percent;
                    db.Cfvo2.Formula = MaximumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Percentile:
                    db.Cfvo2.Type = X14.ConditionalFormattingValueObjectTypeValues.Percentile;
                    db.Cfvo2.Formula = MaximumValue;
                    break;
                case SLConditionalFormatAutoMinMaxValues.Value:
                    db.Cfvo2.Type = X14.ConditionalFormattingValueObjectTypeValues.Max;
                    db.Cfvo2.Formula = string.Empty;
                    break;
            }

            db.FillColor = Color.Clone();
            db.BorderColor = BorderColor.Clone();
            db.NegativeFillColor = NegativeFillColor.Clone();
            db.NegativeBorderColor = NegativeBorderColor.Clone();
            db.AxisColor = AxisColor.Clone();
            db.MinLength = MinLength;
            db.MaxLength = MaxLength;
            db.ShowValue = ShowValue;
            db.Border = Border;
            db.Gradient = Gradient;
            db.Direction = Direction;
            db.NegativeBarColorSameAsPositive = NegativeBarColorSameAsPositive;
            db.NegativeBarBorderColorSameAsPositive = NegativeBarBorderColorSameAsPositive;
            db.AxisPosition = AxisPosition;

            return db;
        }

        internal SLDataBar Clone()
        {
            var db = new SLDataBar();
            db.Is2010 = Is2010;
            db.MinimumType = MinimumType;
            db.MinimumValue = MinimumValue;
            db.MaximumType = MaximumType;
            db.MaximumValue = MaximumValue;
            db.Color = Color.Clone();
            db.BorderColor = BorderColor.Clone();
            db.NegativeFillColor = NegativeFillColor.Clone();
            db.NegativeBorderColor = NegativeBorderColor.Clone();
            db.AxisColor = AxisColor.Clone();
            db.MinLength = MinLength;
            db.MaxLength = MaxLength;
            db.ShowValue = ShowValue;
            db.Border = Border;
            db.Gradient = Gradient;
            db.Direction = Direction;
            db.NegativeBarColorSameAsPositive = NegativeBarColorSameAsPositive;
            db.NegativeBarBorderColorSameAsPositive = NegativeBarBorderColorSameAsPositive;
            db.AxisPosition = AxisPosition;

            return db;
        }
    }
}