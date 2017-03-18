using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.table
{
    internal class SLFilterColumn
    {
        internal bool HasColorFilter;

        internal bool HasCustomFilters;

        internal bool HasDynamicFilter;
        internal bool HasFilters;

        internal bool HasIconFilter;

        internal bool HasTop10;

        internal SLFilterColumn()
        {
            SetAllNull();
        }

        internal SLFilters Filters { get; set; }
        internal SLTop10 Top10 { get; set; }
        internal SLCustomFilters CustomFilters { get; set; }
        internal SLDynamicFilter DynamicFilter { get; set; }
        internal SLColorFilter ColorFilter { get; set; }
        internal SLIconFilter IconFilter { get; set; }

        internal uint ColumnId { get; set; }
        internal bool? HiddenButton { get; set; }
        internal bool? ShowButton { get; set; }

        private void SetAllNull()
        {
            Filters = new SLFilters();
            HasFilters = false;
            Top10 = new SLTop10();
            HasTop10 = false;
            CustomFilters = new SLCustomFilters();
            HasCustomFilters = false;
            DynamicFilter = new SLDynamicFilter();
            HasDynamicFilter = false;
            ColorFilter = new SLColorFilter();
            HasColorFilter = false;
            IconFilter = new SLIconFilter();
            HasIconFilter = false;
            ColumnId = 1;
            HiddenButton = null;
            ShowButton = null;
        }

        private void SetFiltersNull()
        {
            HasFilters = false;
            HasTop10 = false;
            HasCustomFilters = false;
            HasDynamicFilter = false;
            HasColorFilter = false;
            HasIconFilter = false;
        }

        internal void FromFilterColumn(FilterColumn fc)
        {
            SetAllNull();

            if (fc.Filters != null)
            {
                Filters.FromFilters(fc.Filters);
                HasFilters = true;
            }
            if (fc.Top10 != null)
            {
                Top10.FromTop10(fc.Top10);
                HasTop10 = true;
            }
            if (fc.CustomFilters != null)
            {
                CustomFilters.FromCustomFilters(fc.CustomFilters);
                HasCustomFilters = true;
            }
            if (fc.DynamicFilter != null)
            {
                DynamicFilter.FromDynamicFilter(fc.DynamicFilter);
                HasDynamicFilter = true;
            }
            if (fc.ColorFilter != null)
            {
                ColorFilter.FromColorFilter(fc.ColorFilter);
                HasColorFilter = true;
            }
            if (fc.IconFilter != null)
            {
                IconFilter.FromIconFilter(fc.IconFilter);
                HasIconFilter = true;
            }

            ColumnId = fc.ColumnId.Value;
            if ((fc.HiddenButton != null) && fc.HiddenButton.Value) HiddenButton = fc.HiddenButton.Value;
            if ((fc.ShowButton != null) && !fc.ShowButton.Value) ShowButton = fc.ShowButton.Value;
        }

        internal FilterColumn ToFilterColumn()
        {
            var fc = new FilterColumn();

            if (HasFilters) fc.Filters = Filters.ToFilters();
            if (HasTop10) fc.Top10 = Top10.ToTop10();
            if (HasCustomFilters) fc.CustomFilters = CustomFilters.ToCustomFilters();
            if (HasDynamicFilter) fc.DynamicFilter = DynamicFilter.ToDynamicFilter();
            if (HasColorFilter) fc.ColorFilter = ColorFilter.ToColorFilter();
            if (HasIconFilter) fc.IconFilter = IconFilter.ToIconFilter();
            fc.ColumnId = ColumnId;
            if ((HiddenButton != null) && HiddenButton.Value) fc.HiddenButton = HiddenButton.Value;
            if ((ShowButton != null) && !ShowButton.Value) fc.ShowButton = ShowButton.Value;

            return fc;
        }

        internal SLFilterColumn Clone()
        {
            var fc = new SLFilterColumn();
            fc.HasFilters = HasFilters;
            fc.Filters = Filters.Clone();
            fc.HasTop10 = HasTop10;
            fc.Top10 = Top10.Clone();
            fc.HasCustomFilters = HasCustomFilters;
            fc.CustomFilters = CustomFilters.Clone();
            fc.HasDynamicFilter = HasDynamicFilter;
            fc.DynamicFilter = DynamicFilter.Clone();
            fc.HasColorFilter = HasColorFilter;
            fc.ColorFilter = ColorFilter.Clone();
            fc.HasIconFilter = HasIconFilter;
            fc.IconFilter = IconFilter.Clone();
            fc.ColumnId = ColumnId;
            fc.HiddenButton = HiddenButton;
            fc.ShowButton = ShowButton;

            return fc;
        }
    }
}