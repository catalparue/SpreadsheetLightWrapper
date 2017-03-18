using System.Collections.Generic;
using DocumentFormat.OpenXml;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace Ups.Toolkit.SpreadsheetLight.Core.office2010
{
    internal class SLIconSet2010
    {
        // This is true if and only if CustomIcons is used.
        // So we'll just ignore it and focus on the number of CustomIcons instead.
        // internal bool Custom

        internal SLIconSet2010()
        {
            SetAllNull();
        }

        //http://msdn.microsoft.com/en-us/library/documentformat.openxml.office2010.excel.iconset.aspx

        internal List<SLConditionalFormattingValueObject2010> Cfvos { get; set; }
        internal List<SLConditionalFormattingIcon2010> CustomIcons { get; set; }
        internal X14.IconSetTypeValues IconSetType { get; set; }
        internal bool ShowValue { get; set; }
        internal bool Percent { get; set; }
        internal bool Reverse { get; set; }

        private void SetAllNull()
        {
            Cfvos = new List<SLConditionalFormattingValueObject2010>();
            CustomIcons = new List<SLConditionalFormattingIcon2010>();
            IconSetType = X14.IconSetTypeValues.ThreeTrafficLights1;
            ShowValue = true;
            Percent = true;
            Reverse = false;
        }

        internal void FromIconSet(X14.IconSet ics)
        {
            SetAllNull();

            if (ics.IconSetTypes != null) IconSetType = ics.IconSetTypes.Value;
            if (ics.ShowValue != null) ShowValue = ics.ShowValue.Value;
            if (ics.Percent != null) Percent = ics.Percent.Value;
            if (ics.Reverse != null) Reverse = ics.Reverse.Value;

            using (var oxr = OpenXmlReader.Create(ics))
            {
                SLConditionalFormattingValueObject2010 cfvo;
                SLConditionalFormattingIcon2010 cfi;
                while (oxr.Read())
                    if (oxr.ElementType == typeof(X14.ConditionalFormattingValueObject))
                    {
                        cfvo = new SLConditionalFormattingValueObject2010();
                        cfvo.FromConditionalFormattingValueObject(
                            (X14.ConditionalFormattingValueObject) oxr.LoadCurrentElement());
                        Cfvos.Add(cfvo);
                    }
                    else if (oxr.ElementType == typeof(X14.ConditionalFormattingIcon))
                    {
                        cfi = new SLConditionalFormattingIcon2010();
                        cfi.FromConditionalFormattingIcon((X14.ConditionalFormattingIcon) oxr.LoadCurrentElement());
                        CustomIcons.Add(cfi);
                    }
            }
        }

        internal X14.IconSet ToIconSet()
        {
            var ics = new X14.IconSet();
            if (IconSetType != X14.IconSetTypeValues.ThreeTrafficLights1) ics.IconSetTypes = IconSetType;
            if (!ShowValue) ics.ShowValue = ShowValue;
            if (!Percent) ics.Percent = Percent;
            if (Reverse) ics.Reverse = Reverse;
            if (CustomIcons.Count > 0) ics.Custom = true;

            foreach (var cfvo in Cfvos)
                ics.Append(cfvo.ToConditionalFormattingValueObject());

            foreach (var cfi in CustomIcons)
                ics.Append(cfi.ToConditionalFormattingIcon());

            return ics;
        }

        internal SLIconSet2010 Clone()
        {
            var ics = new SLIconSet2010();

            int i;

            ics.Cfvos = new List<SLConditionalFormattingValueObject2010>();
            for (i = 0; i < Cfvos.Count; ++i)
                ics.Cfvos.Add(Cfvos[i].Clone());

            ics.CustomIcons = new List<SLConditionalFormattingIcon2010>();
            for (i = 0; i < CustomIcons.Count; ++i)
                ics.CustomIcons.Add(CustomIcons[i].Clone());

            ics.IconSetType = IconSetType;
            ics.ShowValue = ShowValue;
            ics.Percent = Percent;
            ics.Reverse = Reverse;

            return ics;
        }
    }
}