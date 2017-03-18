using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.table
{
    internal class SLCustomFilters
    {
        internal bool HasFirstOperator;

        internal bool HasSecondOperator;
        internal bool OneCustomFilter;
        private FilterOperatorValues vFirstOperator;
        private FilterOperatorValues vSecondOperator;

        internal SLCustomFilters()
        {
            var cf = new CustomFilter();
            cf.Operator = FilterOperatorValues.Equal;
            cf.Val = "";
        }

        internal FilterOperatorValues FirstOperator
        {
            get { return vFirstOperator; }
            set
            {
                vFirstOperator = value;
                HasFirstOperator = vFirstOperator != FilterOperatorValues.Equal ? true : false;
            }
        }

        internal string FirstVal { get; set; }

        internal FilterOperatorValues SecondOperator
        {
            get { return vSecondOperator; }
            set
            {
                vSecondOperator = value;
                HasSecondOperator = vSecondOperator != FilterOperatorValues.Equal ? true : false;
            }
        }

        internal string SecondVal { get; set; }

        internal bool? And { get; set; }

        private void SetAllNull()
        {
            OneCustomFilter = true;
            vFirstOperator = FilterOperatorValues.Equal;
            HasFirstOperator = false;
            FirstVal = string.Empty;
            vSecondOperator = FilterOperatorValues.Equal;
            HasSecondOperator = false;
            SecondVal = string.Empty;
            And = null;
        }

        internal void FromCustomFilters(CustomFilters cfs)
        {
            SetAllNull();

            if ((cfs.And != null) && cfs.And.Value) And = cfs.And.Value;

            var i = 0;
            CustomFilter cf;
            using (var oxr = OpenXmlReader.Create(cfs))
            {
                while (oxr.Read())
                    if (oxr.ElementType == typeof(CustomFilter))
                    {
                        ++i;
                        cf = (CustomFilter) oxr.LoadCurrentElement();
                        if (i == 1)
                        {
                            OneCustomFilter = true;
                            if (cf.Operator != null) FirstOperator = cf.Operator.Value;
                            if (cf.Val != null) FirstVal = cf.Val.Value;
                        }
                        else if (i == 2)
                        {
                            OneCustomFilter = false;
                            if (cf.Operator != null) SecondOperator = cf.Operator.Value;
                            if (cf.Val != null) SecondVal = cf.Val.Value;
                        }
                        else
                        {
                            break;
                        }
                    }
            }
        }

        internal CustomFilters ToCustomFilters()
        {
            var cfs = new CustomFilters();
            if ((And != null) && And.Value) cfs.And = And.Value;

            CustomFilter cf;
            if (OneCustomFilter)
            {
                cf = new CustomFilter();
                if (HasFirstOperator) cf.Operator = FirstOperator;
                cf.Val = FirstVal;
                cfs.Append(cf);
            }
            else
            {
                cf = new CustomFilter();
                if (HasFirstOperator) cf.Operator = FirstOperator;
                cf.Val = FirstVal;
                cfs.Append(cf);

                cf = new CustomFilter();
                if (HasSecondOperator) cf.Operator = SecondOperator;
                cf.Val = SecondVal;
                cfs.Append(cf);
            }

            return cfs;
        }

        internal SLCustomFilters Clone()
        {
            var cfs = new SLCustomFilters();
            cfs.OneCustomFilter = OneCustomFilter;
            cfs.HasFirstOperator = HasFirstOperator;
            cfs.vFirstOperator = vFirstOperator;
            cfs.FirstVal = FirstVal;
            cfs.HasSecondOperator = HasSecondOperator;
            cfs.vSecondOperator = vSecondOperator;
            cfs.SecondVal = SecondVal;
            cfs.And = And;

            return cfs;
        }
    }
}