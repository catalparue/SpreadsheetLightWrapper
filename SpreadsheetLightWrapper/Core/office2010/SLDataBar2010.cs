using System.Collections.Generic;
using DocumentFormat.OpenXml;
using SpreadsheetLightWrapper.Core.style;
using Color = System.Drawing.Color;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace SpreadsheetLightWrapper.Core.office2010
{
    internal class SLDataBar2010
    {
        internal SLDataBar2010()
        {
            SetAllNull();
        }

        //http://msdn.microsoft.com/en-us/library/documentformat.openxml.office2010.excel.databar.aspx

        internal SLConditionalFormattingValueObject2010 Cfvo1 { get; set; }
        internal SLConditionalFormattingValueObject2010 Cfvo2 { get; set; }
        internal SLColor FillColor { get; set; }
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
            Cfvo1 = new SLConditionalFormattingValueObject2010();
            Cfvo2 = new SLConditionalFormattingValueObject2010();
            FillColor = new SLColor(new List<Color>(), new List<Color>());
            BorderColor = new SLColor(new List<Color>(), new List<Color>());
            NegativeFillColor = new SLColor(new List<Color>(), new List<Color>());
            NegativeBorderColor = new SLColor(new List<Color>(), new List<Color>());
            AxisColor = new SLColor(new List<Color>(), new List<Color>());

            MinLength = 10;
            MaxLength = 90;
            ShowValue = true;
            Border = false;
            Gradient = true;
            Direction = X14.DataBarDirectionValues.Context;
            NegativeBarColorSameAsPositive = false;
            NegativeBarBorderColorSameAsPositive = true;
            AxisPosition = X14.DataBarAxisPositionValues.Automatic;
        }

        internal void FromDataBar(X14.DataBar db)
        {
            SetAllNull();

            using (var oxr = OpenXmlReader.Create(db))
            {
                var i = 0;
                while (oxr.Read())
                    if (oxr.ElementType == typeof(X14.ConditionalFormattingValueObject))
                    {
                        if (i == 0)
                        {
                            Cfvo1.FromConditionalFormattingValueObject(
                                (X14.ConditionalFormattingValueObject) oxr.LoadCurrentElement());
                            ++i;
                        }
                        else if (i == 1)
                        {
                            Cfvo2.FromConditionalFormattingValueObject(
                                (X14.ConditionalFormattingValueObject) oxr.LoadCurrentElement());
                            ++i;
                        }
                    }
                    else if (oxr.ElementType == typeof(X14.FillColor))
                    {
                        FillColor.FromFillColor((X14.FillColor) oxr.LoadCurrentElement());
                    }
                    else if (oxr.ElementType == typeof(X14.BorderColor))
                    {
                        BorderColor.FromBorderColor((X14.BorderColor) oxr.LoadCurrentElement());
                    }
                    else if (oxr.ElementType == typeof(X14.NegativeFillColor))
                    {
                        NegativeFillColor.FromNegativeFillColor((X14.NegativeFillColor) oxr.LoadCurrentElement());
                    }
                    else if (oxr.ElementType == typeof(X14.NegativeBorderColor))
                    {
                        NegativeBorderColor.FromNegativeBorderColor((X14.NegativeBorderColor) oxr.LoadCurrentElement());
                    }
                    else if (oxr.ElementType == typeof(X14.BarAxisColor))
                    {
                        AxisColor.FromBarAxisColor((X14.BarAxisColor) oxr.LoadCurrentElement());
                    }
            }

            if (db.MinLength != null) MinLength = db.MinLength.Value;
            if (db.MaxLength != null) MaxLength = db.MaxLength.Value;
            if (db.ShowValue != null) ShowValue = db.ShowValue.Value;
            if (db.Border != null) Border = db.Border.Value;
            if (db.Gradient != null) Gradient = db.Gradient.Value;
            if (db.Direction != null) Direction = db.Direction.Value;
            if (db.NegativeBarColorSameAsPositive != null)
                NegativeBarColorSameAsPositive = db.NegativeBarColorSameAsPositive.Value;
            if (db.NegativeBarBorderColorSameAsPositive != null)
                NegativeBarBorderColorSameAsPositive = db.NegativeBarBorderColorSameAsPositive.Value;
            if (db.AxisPosition != null) AxisPosition = db.AxisPosition.Value;
        }

        internal X14.DataBar ToDataBar(bool RenderFillColor)
        {
            var db = new X14.DataBar();
            if (MinLength != 10) db.MinLength = MinLength;

            // according to Open XML specs, this cannot be more than 100 percent.
            if (MaxLength > 100) MaxLength = 100;
            if (MaxLength != 90) db.MaxLength = MaxLength;

            if (!ShowValue) db.ShowValue = ShowValue;
            if (Border) db.Border = Border;
            if (!Gradient) db.Gradient = Gradient;
            if (Direction != X14.DataBarDirectionValues.Context) db.Direction = Direction;
            if (NegativeBarColorSameAsPositive) db.NegativeBarColorSameAsPositive = NegativeBarColorSameAsPositive;
            if (!NegativeBarBorderColorSameAsPositive)
                db.NegativeBarBorderColorSameAsPositive = NegativeBarBorderColorSameAsPositive;
            if (AxisPosition != X14.DataBarAxisPositionValues.Automatic) db.AxisPosition = AxisPosition;

            db.Append(Cfvo1.ToConditionalFormattingValueObject());
            db.Append(Cfvo2.ToConditionalFormattingValueObject());

            // The condition is mainly if the priority of the parent rule exists. See Open XML specs.
            if (RenderFillColor) db.Append(FillColor.ToFillColor());

            if (Border) db.Append(BorderColor.ToBorderColor());
            if (!NegativeBarColorSameAsPositive) db.Append(NegativeFillColor.ToNegativeFillColor());
            if (!NegativeBarBorderColorSameAsPositive && Border) db.Append(NegativeBorderColor.ToNegativeBorderColor());
            if (AxisPosition != X14.DataBarAxisPositionValues.None) db.Append(AxisColor.ToBarAxisColor());

            return db;
        }

        internal SLDataBar2010 Clone()
        {
            var db = new SLDataBar2010();
            db.Cfvo1 = Cfvo1.Clone();
            db.Cfvo2 = Cfvo2.Clone();
            db.FillColor = FillColor.Clone();
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