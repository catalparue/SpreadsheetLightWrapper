using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.pivottable
{
    internal class SLConditionalFormat
    {
        internal SLConditionalFormat()
        {
            SetAllNull();
        }

        internal List<SLPivotArea> PivotAreas { get; set; }
        internal ScopeValues Scope { get; set; }
        internal RuleValues Type { get; set; }
        internal uint Priority { get; set; }

        private void SetAllNull()
        {
            PivotAreas = new List<SLPivotArea>();
            Scope = ScopeValues.Selection;
            Type = RuleValues.None;
            Priority = 0;
        }

        internal void FromConditionalFormat(ConditionalFormat cf)
        {
            SetAllNull();

            if (cf.Scope != null) Scope = cf.Scope.Value;
            if (cf.Type != null) Type = cf.Type.Value;
            if (cf.Priority != null) Priority = cf.Priority.Value;

            SLPivotArea pa;
            using (var oxr = OpenXmlReader.Create(cf))
            {
                while (oxr.Read())
                    if (oxr.ElementType == typeof(PivotArea))
                    {
                        pa = new SLPivotArea();
                        pa.FromPivotArea((PivotArea) oxr.LoadCurrentElement());
                        PivotAreas.Add(pa);
                    }
            }
        }

        internal ConditionalFormat ToConditionalFormat()
        {
            var cf = new ConditionalFormat();
            cf.PivotAreas = new PivotAreas {Count = (uint) PivotAreas.Count};
            foreach (var pa in PivotAreas)
                cf.PivotAreas.Append(pa.ToPivotArea());

            if (Scope != ScopeValues.Selection) cf.Scope = Scope;
            if (Type != RuleValues.None) cf.Type = Type;
            cf.Priority = Priority;

            return cf;
        }

        internal SLConditionalFormat Clone()
        {
            var cf = new SLConditionalFormat();
            cf.Scope = Scope;
            cf.Type = Type;
            cf.Priority = Priority;

            cf.PivotAreas = new List<SLPivotArea>();
            foreach (var pa in PivotAreas)
                cf.PivotAreas.Add(pa.Clone());

            return cf;
        }
    }
}