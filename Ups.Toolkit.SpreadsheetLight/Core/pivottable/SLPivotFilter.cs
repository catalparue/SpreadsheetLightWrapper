using DocumentFormat.OpenXml.Spreadsheet;
using Ups.Toolkit.SpreadsheetLight.Core.misc;

namespace Ups.Toolkit.SpreadsheetLight.Core.pivottable
{
    internal class SLPivotFilter
    {
        internal SLPivotFilter()
        {
            SetAllNull();
        }

        internal SLAutoFilter AutoFilter { get; set; }

        internal uint Field { get; set; }
        internal uint? MemberPropertyFieldId { get; set; }
        internal PivotFilterValues Type { get; set; }
        internal int EvaluationOrder { get; set; }
        internal uint Id { get; set; }
        internal uint? MeasureHierarchy { get; set; }
        internal uint? MeasureField { get; set; }
        internal string Name { get; set; }
        internal string Description { get; set; }
        internal string StringValue1 { get; set; }
        internal string StringValue2 { get; set; }

        private void SetAllNull()
        {
            AutoFilter = new SLAutoFilter();

            Field = 0;
            MemberPropertyFieldId = null;
            Type = PivotFilterValues.Unknown;
            EvaluationOrder = 0;
            Id = 0;
            MeasureHierarchy = null;
            MeasureField = null;
            Name = "";
            Description = "";
            StringValue1 = "";
            StringValue2 = "";
        }

        internal void FromPivotFilter(PivotFilter pf)
        {
            SetAllNull();

            if (pf.Field != null) Field = pf.Field.Value;
            if (pf.MemberPropertyFieldId != null) MemberPropertyFieldId = pf.MemberPropertyFieldId.Value;
            if (pf.Type != null) Type = pf.Type.Value;
            if (pf.EvaluationOrder != null) EvaluationOrder = pf.EvaluationOrder.Value;
            if (pf.Id != null) Id = pf.Id.Value;
            if (pf.MeasureHierarchy != null) MeasureHierarchy = pf.MeasureHierarchy.Value;
            if (pf.MeasureField != null) MeasureField = pf.MeasureField.Value;
            if (pf.Name != null) Name = pf.Name.Value;
            if (pf.Description != null) Description = pf.Description.Value;
            if (pf.StringValue1 != null) StringValue1 = pf.StringValue1.Value;
            if (pf.StringValue2 != null) StringValue2 = pf.StringValue2.Value;

            if (pf.AutoFilter != null) AutoFilter.FromAutoFilter(pf.AutoFilter);
        }

        internal PivotFilter ToPivotFilter()
        {
            var pf = new PivotFilter();
            pf.Field = Field;
            if (MemberPropertyFieldId != null) pf.MemberPropertyFieldId = MemberPropertyFieldId.Value;
            pf.Type = Type;
            if (EvaluationOrder != 0) pf.EvaluationOrder = EvaluationOrder;
            pf.Id = Id;
            if (MeasureHierarchy != null) pf.MeasureHierarchy = MeasureHierarchy.Value;
            if (MeasureField != null) pf.MeasureField = MeasureField.Value;
            if ((Name != null) && (Name.Length > 0)) pf.Name = Name;
            if ((Description != null) && (Description.Length > 0)) pf.Description = Description;
            if ((StringValue1 != null) && (StringValue1.Length > 0)) pf.StringValue1 = StringValue1;
            if ((StringValue2 != null) && (StringValue2.Length > 0)) pf.StringValue2 = StringValue2;

            pf.AutoFilter = AutoFilter.ToAutoFilter();

            return pf;
        }

        internal SLPivotFilter Clone()
        {
            var pf = new SLPivotFilter();
            pf.Field = Field;
            pf.MemberPropertyFieldId = MemberPropertyFieldId;
            pf.Type = Type;
            pf.EvaluationOrder = EvaluationOrder;
            pf.Id = Id;
            pf.MeasureHierarchy = MeasureHierarchy;
            pf.MeasureField = MeasureField;
            pf.Name = Name;
            pf.Description = Description;
            pf.StringValue1 = StringValue1;
            pf.StringValue2 = StringValue2;

            pf.AutoFilter = AutoFilter.Clone();

            return pf;
        }
    }
}