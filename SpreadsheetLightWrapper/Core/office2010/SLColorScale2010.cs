using System.Collections.Generic;
using DocumentFormat.OpenXml;
using SpreadsheetLightWrapper.Core.style;
using Color = System.Drawing.Color;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace SpreadsheetLightWrapper.Core.office2010
{
    internal class SLColorScale2010
    {
        internal SLColorScale2010()
        {
            SetAllNull();
        }

        //http://msdn.microsoft.com/en-us/library/documentformat.openxml.office2010.excel.colorscale.aspx

        internal List<SLConditionalFormattingValueObject2010> Cfvos { get; set; }
        internal List<SLColor> Colors { get; set; }

        private void SetAllNull()
        {
            Cfvos = new List<SLConditionalFormattingValueObject2010>();
            Colors = new List<SLColor>();
        }

        internal void FromColorScale(X14.ColorScale cs)
        {
            SetAllNull();

            SLConditionalFormattingValueObject2010 cfvo;
            SLColor clr;
            using (var oxr = OpenXmlReader.Create(cs))
            {
                while (oxr.Read())
                    if (oxr.ElementType == typeof(X14.ConditionalFormattingValueObject))
                    {
                        cfvo = new SLConditionalFormattingValueObject2010();
                        cfvo.FromConditionalFormattingValueObject(
                            (X14.ConditionalFormattingValueObject) oxr.LoadCurrentElement());
                        Cfvos.Add(cfvo);
                    }
                    else if (oxr.ElementType == typeof(X14.Color))
                    {
                        clr = new SLColor(new List<Color>(), new List<Color>());
                        clr.FromExcel2010Color((X14.Color) oxr.LoadCurrentElement());
                        Colors.Add(clr);
                    }
            }
        }

        internal X14.ColorScale ToColorScale()
        {
            var cs = new X14.ColorScale();
            foreach (var cfvo in Cfvos)
                cs.Append(cfvo.ToConditionalFormattingValueObject());
            foreach (var clr in Colors)
                cs.Append(clr.ToExcel2010Color());

            return cs;
        }

        internal SLColorScale2010 Clone()
        {
            var cs = new SLColorScale2010();

            int i;
            cs.Cfvos = new List<SLConditionalFormattingValueObject2010>();
            for (i = 0; i < Cfvos.Count; ++i)
                cs.Cfvos.Add(Cfvos[i].Clone());

            cs.Colors = new List<SLColor>();
            for (i = 0; i < Colors.Count; ++i)
                cs.Colors.Add(Colors[i].Clone());

            return cs;
        }
    }
}