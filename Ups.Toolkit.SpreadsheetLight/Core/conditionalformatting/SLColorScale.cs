using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Ups.Toolkit.SpreadsheetLight.Core.office2010;
using Ups.Toolkit.SpreadsheetLight.Core.style;

namespace Ups.Toolkit.SpreadsheetLight.Core.conditionalformatting
{
    internal class SLColorScale
    {
        internal SLColorScale()
        {
            SetAllNull();
        }

        internal List<SLConditionalFormatValueObject> Cfvos { get; set; }
        internal List<SLColor> Colors { get; set; }

        private void SetAllNull()
        {
            Cfvos = new List<SLConditionalFormatValueObject>();
            Colors = new List<SLColor>();
        }

        internal void FromColorScale(ColorScale cs)
        {
            SetAllNull();

            SLConditionalFormatValueObject cfvo;
            SLColor clr;
            using (var oxr = OpenXmlReader.Create(cs))
            {
                while (oxr.Read())
                    if (oxr.ElementType == typeof(ConditionalFormatValueObject))
                    {
                        cfvo = new SLConditionalFormatValueObject();
                        cfvo.FromConditionalFormatValueObject((ConditionalFormatValueObject) oxr.LoadCurrentElement());
                        Cfvos.Add(cfvo);
                    }
                    else if (oxr.ElementType == typeof(Color))
                    {
                        clr = new SLColor(new List<System.Drawing.Color>(), new List<System.Drawing.Color>());
                        clr.FromSpreadsheetColor((Color) oxr.LoadCurrentElement());
                        Colors.Add(clr);
                    }
            }
        }

        internal ColorScale ToColorScale()
        {
            var cs = new ColorScale();
            foreach (var cfvo in Cfvos)
                cs.Append(cfvo.ToConditionalFormatValueObject());
            foreach (var clr in Colors)
                cs.Append(clr.ToSpreadsheetColor());

            return cs;
        }

        internal SLColorScale2010 ToSLColorScale2010()
        {
            var cs2010 = new SLColorScale2010();
            foreach (var cfvo in Cfvos)
                cs2010.Cfvos.Add(cfvo.ToSLConditionalFormattingValueObject2010());
            foreach (var clr in Colors)
                cs2010.Colors.Add(clr.Clone());

            return cs2010;
        }

        internal SLColorScale Clone()
        {
            var cs = new SLColorScale();

            int i;
            cs.Cfvos = new List<SLConditionalFormatValueObject>();
            for (i = 0; i < Cfvos.Count; ++i)
                cs.Cfvos.Add(Cfvos[i].Clone());

            cs.Colors = new List<SLColor>();
            for (i = 0; i < Colors.Count; ++i)
                cs.Colors.Add(Colors[i].Clone());

            return cs;
        }
    }
}