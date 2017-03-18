using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.pivottable
{
    internal class SLPivotArea
    {
        internal SLPivotArea()
        {
            SetAllNull();
        }

        private List<SLPivotAreaReference> PivotAreaReferences { get; set; }

        internal int? Field { get; set; }
        internal PivotAreaValues Type { get; set; }
        internal bool DataOnly { get; set; }
        internal bool LabelOnly { get; set; }
        internal bool GrandRow { get; set; }
        internal bool GrandColumn { get; set; }
        internal bool CacheIndex { get; set; }
        internal bool Outline { get; set; }
        internal string Offset { get; set; }
        internal bool CollapsedLevelsAreSubtotals { get; set; }
        internal PivotTableAxisValues? Axis { get; set; }
        internal uint? FieldPosition { get; set; }

        private void SetAllNull()
        {
            PivotAreaReferences = new List<SLPivotAreaReference>();

            Field = null;
            Type = PivotAreaValues.Normal;
            DataOnly = true;
            LabelOnly = false;
            GrandRow = false;
            GrandColumn = false;
            CacheIndex = false;
            Outline = true;
            Offset = "";
            CollapsedLevelsAreSubtotals = false;
            Axis = null;
            FieldPosition = null;
        }

        internal void FromPivotArea(PivotArea pa)
        {
            SetAllNull();

            if (pa.Field != null) Field = pa.Field.Value;
            if (pa.Type != null) Type = pa.Type.Value;
            if (pa.DataOnly != null) DataOnly = pa.DataOnly.Value;
            if (pa.LabelOnly != null) LabelOnly = pa.LabelOnly.Value;
            if (pa.GrandRow != null) GrandRow = pa.GrandRow.Value;
            if (pa.GrandColumn != null) GrandColumn = pa.GrandColumn.Value;
            if (pa.CacheIndex != null) CacheIndex = pa.CacheIndex.Value;
            if (pa.Outline != null) Outline = pa.Outline.Value;
            if (pa.Offset != null) Offset = pa.Offset.Value;
            if (pa.CollapsedLevelsAreSubtotals != null)
                CollapsedLevelsAreSubtotals = pa.CollapsedLevelsAreSubtotals.Value;
            if (pa.Axis != null) Axis = pa.Axis.Value;
            if (pa.FieldPosition != null) FieldPosition = pa.FieldPosition.Value;

            SLPivotAreaReference par;
            using (var oxr = OpenXmlReader.Create(pa))
            {
                while (oxr.Read())
                    if (oxr.ElementType == typeof(PivotAreaReference))
                    {
                        par = new SLPivotAreaReference();
                        par.FromPivotAreaReference((PivotAreaReference) oxr.LoadCurrentElement());
                        PivotAreaReferences.Add(par);
                    }
            }
        }

        internal PivotArea ToPivotArea()
        {
            var pa = new PivotArea();
            if (Field != null) pa.Field = Field.Value;
            if (Type != PivotAreaValues.Normal) pa.Type = Type;
            if (DataOnly != true) pa.DataOnly = DataOnly;
            if (LabelOnly) pa.LabelOnly = LabelOnly;
            if (GrandRow) pa.GrandRow = GrandRow;
            if (GrandColumn) pa.GrandColumn = GrandColumn;
            if (CacheIndex) pa.CacheIndex = CacheIndex;
            if (Outline != true) pa.Outline = Outline;
            if ((Offset != null) && (Offset.Length > 0)) pa.Offset = Offset;
            if (CollapsedLevelsAreSubtotals) pa.CollapsedLevelsAreSubtotals = CollapsedLevelsAreSubtotals;
            if (Axis != null) pa.Axis = Axis.Value;
            if (FieldPosition != null) pa.FieldPosition = FieldPosition.Value;

            if (PivotAreaReferences.Count > 0)
            {
                pa.PivotAreaReferences = new PivotAreaReferences();
                foreach (var par in PivotAreaReferences)
                    pa.PivotAreaReferences.Append(par.ToPivotAreaReference());
            }

            return pa;
        }

        internal SLPivotArea Clone()
        {
            var pa = new SLPivotArea();
            pa.Field = Field;
            pa.Type = Type;
            pa.DataOnly = DataOnly;
            pa.LabelOnly = LabelOnly;
            pa.GrandRow = GrandRow;
            pa.GrandColumn = GrandColumn;
            pa.CacheIndex = CacheIndex;
            pa.Outline = Outline;
            pa.Offset = Offset;
            pa.CollapsedLevelsAreSubtotals = CollapsedLevelsAreSubtotals;
            pa.Axis = Axis;
            pa.FieldPosition = FieldPosition;

            pa.PivotAreaReferences = new List<SLPivotAreaReference>();
            foreach (var par in PivotAreaReferences)
                pa.PivotAreaReferences.Add(par.Clone());

            return pa;
        }
    }
}