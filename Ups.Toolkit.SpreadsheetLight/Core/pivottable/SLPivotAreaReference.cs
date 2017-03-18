using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.pivottable
{
    internal class SLPivotAreaReference
    {
        internal SLPivotAreaReference()
        {
            SetAllNull();
        }

        internal List<uint> FieldItems { get; set; }

        internal uint? Field { get; set; }
        //internal uint Count { get; set; }
        internal bool Selected { get; set; }
        internal bool ByPosition { get; set; }
        internal bool Relative { get; set; }
        internal bool DefaultSubtotal { get; set; }
        internal bool SumSubtotal { get; set; }
        internal bool CountASubtotal { get; set; }
        internal bool AverageSubtotal { get; set; }
        internal bool MaxSubtotal { get; set; }
        internal bool MinSubtotal { get; set; }
        internal bool ApplyProductInSubtotal { get; set; }
        internal bool CountSubtotal { get; set; }
        internal bool ApplyStandardDeviationInSubtotal { get; set; }
        internal bool ApplyStandardDeviationPInSubtotal { get; set; }
        internal bool ApplyVarianceInSubtotal { get; set; }
        internal bool ApplyVariancePInSubtotal { get; set; }

        private void SetAllNull()
        {
            FieldItems = new List<uint>();
            Field = null;
            //this.Count = 0;
            Selected = true;
            ByPosition = false;
            Relative = false;
            DefaultSubtotal = false;
            SumSubtotal = false;
            CountASubtotal = false;
            AverageSubtotal = false;
            MaxSubtotal = false;
            MinSubtotal = false;
            ApplyProductInSubtotal = false;
            CountSubtotal = false;
            ApplyStandardDeviationInSubtotal = false;
            ApplyStandardDeviationPInSubtotal = false;
            ApplyVarianceInSubtotal = false;
            ApplyVariancePInSubtotal = false;
        }

        internal void FromPivotAreaReference(PivotAreaReference par)
        {
            SetAllNull();

            if (par.Field != null) Field = par.Field.Value;
            if (par.Selected != null) Selected = par.Selected.Value;
            if (par.ByPosition != null) ByPosition = par.ByPosition.Value;
            if (par.Relative != null) Relative = par.Relative.Value;
            if (par.DefaultSubtotal != null) DefaultSubtotal = par.DefaultSubtotal.Value;
            if (par.SumSubtotal != null) SumSubtotal = par.SumSubtotal.Value;
            if (par.CountASubtotal != null) CountASubtotal = par.CountASubtotal.Value;
            if (par.AverageSubtotal != null) AverageSubtotal = par.AverageSubtotal.Value;
            if (par.MaxSubtotal != null) MaxSubtotal = par.MaxSubtotal.Value;
            if (par.MinSubtotal != null) MinSubtotal = par.MinSubtotal.Value;
            if (par.ApplyProductInSubtotal != null) ApplyProductInSubtotal = par.ApplyProductInSubtotal.Value;
            if (par.CountSubtotal != null) CountSubtotal = par.CountSubtotal.Value;
            if (par.ApplyStandardDeviationInSubtotal != null)
                ApplyStandardDeviationInSubtotal = par.ApplyStandardDeviationInSubtotal.Value;
            if (par.ApplyStandardDeviationPInSubtotal != null)
                ApplyStandardDeviationPInSubtotal = par.ApplyStandardDeviationPInSubtotal.Value;
            if (par.ApplyVarianceInSubtotal != null) ApplyVarianceInSubtotal = par.ApplyVarianceInSubtotal.Value;
            if (par.ApplyVariancePInSubtotal != null) ApplyVariancePInSubtotal = par.ApplyVariancePInSubtotal.Value;

            FieldItem fi;
            using (var oxr = OpenXmlReader.Create(par))
            {
                while (oxr.Read())
                    if (oxr.ElementType == typeof(FieldItem))
                    {
                        fi = (FieldItem) oxr.LoadCurrentElement();
                        // the Val property is required
                        FieldItems.Add(fi.Val.Value);
                    }
            }
        }

        internal PivotAreaReference ToPivotAreaReference()
        {
            var par = new PivotAreaReference();
            if (Field != null) par.Field = Field.Value;
            par.Count = (uint) FieldItems.Count;
            if (Selected != true) par.Selected = Selected;
            if (ByPosition) par.ByPosition = ByPosition;
            if (Relative) par.Relative = Relative;
            if (DefaultSubtotal) par.DefaultSubtotal = DefaultSubtotal;
            if (SumSubtotal) par.SumSubtotal = SumSubtotal;
            if (CountASubtotal) par.CountASubtotal = CountASubtotal;
            if (AverageSubtotal) par.AverageSubtotal = AverageSubtotal;
            if (MaxSubtotal) par.MaxSubtotal = MaxSubtotal;
            if (MinSubtotal) par.MinSubtotal = MinSubtotal;
            if (ApplyProductInSubtotal) par.ApplyProductInSubtotal = ApplyProductInSubtotal;
            if (CountSubtotal) par.CountSubtotal = CountSubtotal;
            if (ApplyStandardDeviationInSubtotal)
                par.ApplyStandardDeviationInSubtotal = ApplyStandardDeviationInSubtotal;
            if (ApplyStandardDeviationPInSubtotal)
                par.ApplyStandardDeviationPInSubtotal = ApplyStandardDeviationPInSubtotal;
            if (ApplyVarianceInSubtotal) par.ApplyVarianceInSubtotal = ApplyVarianceInSubtotal;
            if (ApplyVariancePInSubtotal) par.ApplyVariancePInSubtotal = ApplyVariancePInSubtotal;

            foreach (var i in FieldItems)
                par.Append(new FieldItem {Val = i});

            return par;
        }

        internal SLPivotAreaReference Clone()
        {
            var par = new SLPivotAreaReference();
            par.Field = Field;
            par.Selected = Selected;
            par.ByPosition = ByPosition;
            par.Relative = Relative;
            par.DefaultSubtotal = DefaultSubtotal;
            par.SumSubtotal = SumSubtotal;
            par.CountASubtotal = CountASubtotal;
            par.AverageSubtotal = AverageSubtotal;
            par.MaxSubtotal = MaxSubtotal;
            par.MinSubtotal = MinSubtotal;
            par.ApplyProductInSubtotal = ApplyProductInSubtotal;
            par.CountSubtotal = CountSubtotal;
            par.ApplyStandardDeviationInSubtotal = ApplyStandardDeviationInSubtotal;
            par.ApplyStandardDeviationPInSubtotal = ApplyStandardDeviationPInSubtotal;
            par.ApplyVarianceInSubtotal = ApplyVarianceInSubtotal;
            par.ApplyVariancePInSubtotal = ApplyVariancePInSubtotal;

            par.FieldItems = new List<uint>();
            foreach (var i in FieldItems)
                par.FieldItems.Add(i);

            return par;
        }
    }
}