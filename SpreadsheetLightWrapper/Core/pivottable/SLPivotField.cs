using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.pivottable
{
    internal class SLPivotField
    {
        internal bool HasAutoSortScope;

        internal SLPivotField()
        {
            SetAllNull();
        }

        internal List<SLItem> Items { get; set; }
        internal SLAutoSortScope AutoSortScope { get; set; }

        internal string Name { get; set; }
        internal PivotTableAxisValues? Axis { get; set; }
        internal bool DataField { get; set; }
        internal string SubtotalCaption { get; set; }
        internal bool ShowDropDowns { get; set; }
        internal bool HiddenLevel { get; set; }
        internal string UniqueMemberProperty { get; set; }
        internal bool Compact { get; set; }
        internal bool AllDrilled { get; set; }
        internal uint? NumberFormatId { get; set; }
        internal bool Outline { get; set; }
        internal bool SubtotalTop { get; set; }
        internal bool DragToRow { get; set; }
        internal bool DragToColumn { get; set; }
        internal bool MultipleItemSelectionAllowed { get; set; }
        internal bool DragToPage { get; set; }
        internal bool DragToData { get; set; }
        internal bool DragOff { get; set; }
        internal bool ShowAll { get; set; }
        internal bool InsertBlankRow { get; set; }
        internal bool ServerField { get; set; }
        internal bool InsertPageBreak { get; set; }
        internal bool AutoShow { get; set; }
        internal bool TopAutoShow { get; set; }
        internal bool HideNewItems { get; set; }
        internal bool MeasureFilter { get; set; }
        internal bool IncludeNewItemsInFilter { get; set; }
        internal uint ItemPageCount { get; set; }
        internal FieldSortValues SortType { get; set; }
        internal bool? DataSourceSort { get; set; }
        internal bool NonAutoSortDefault { get; set; }
        internal uint? RankBy { get; set; }
        internal bool DefaultSubtotal { get; set; }
        internal bool SumSubtotal { get; set; }
        internal bool CountASubtotal { get; set; }
        internal bool AverageSubTotal { get; set; }
        internal bool MaxSubtotal { get; set; }
        internal bool MinSubtotal { get; set; }
        internal bool ApplyProductInSubtotal { get; set; }
        internal bool CountSubtotal { get; set; }
        internal bool ApplyStandardDeviationInSubtotal { get; set; }
        internal bool ApplyStandardDeviationPInSubtotal { get; set; }
        internal bool ApplyVarianceInSubtotal { get; set; }
        internal bool ApplyVariancePInSubtotal { get; set; }
        internal bool ShowPropCell { get; set; }
        internal bool ShowPropertyTooltip { get; set; }
        internal bool ShowPropAsCaption { get; set; }
        internal bool DefaultAttributeDrillState { get; set; }

        private void SetAllNull()
        {
            Items = new List<SLItem>();

            AutoSortScope = new SLAutoSortScope();
            HasAutoSortScope = false;

            Name = "";
            Axis = null;
            DataField = false;
            SubtotalCaption = "";
            ShowDropDowns = true;
            HiddenLevel = false;
            UniqueMemberProperty = "";
            Compact = true;
            AllDrilled = false;
            NumberFormatId = null;
            Outline = true;
            SubtotalTop = true;
            DragToRow = true;
            DragToColumn = true;
            MultipleItemSelectionAllowed = false;
            DragToPage = true;
            DragToData = true;
            DragOff = true;
            ShowAll = true;
            InsertBlankRow = false;
            ServerField = false;
            InsertPageBreak = false;
            AutoShow = false;
            TopAutoShow = true;
            HideNewItems = false;
            MeasureFilter = false;
            IncludeNewItemsInFilter = false;
            ItemPageCount = 10;
            SortType = FieldSortValues.Manual;
            DataSourceSort = null;
            NonAutoSortDefault = false;
            RankBy = null;
            DefaultSubtotal = true;
            SumSubtotal = false;
            CountASubtotal = false;
            AverageSubTotal = false;
            MaxSubtotal = false;
            MinSubtotal = false;
            ApplyProductInSubtotal = false;
            CountSubtotal = false;
            ApplyStandardDeviationInSubtotal = false;
            ApplyStandardDeviationPInSubtotal = false;
            ApplyVarianceInSubtotal = false;
            ApplyVariancePInSubtotal = false;
            ShowPropCell = false;
            ShowPropertyTooltip = false;
            ShowPropAsCaption = false;
            DefaultAttributeDrillState = false;
        }

        internal void FromPivotField(PivotField pf)
        {
            SetAllNull();

            if (pf.Name != null) Name = pf.Name.Value;
            if (pf.Axis != null) Axis = pf.Axis.Value;
            if (pf.DataField != null) DataField = pf.DataField.Value;
            if (pf.SubtotalCaption != null) SubtotalCaption = pf.SubtotalCaption.Value;
            if (pf.ShowDropDowns != null) ShowDropDowns = pf.ShowDropDowns.Value;
            if (pf.HiddenLevel != null) HiddenLevel = pf.HiddenLevel.Value;
            if (pf.UniqueMemberProperty != null) UniqueMemberProperty = pf.UniqueMemberProperty.Value;
            if (pf.Compact != null) Compact = pf.Compact.Value;
            if (pf.AllDrilled != null) AllDrilled = pf.AllDrilled.Value;
            if (pf.NumberFormatId != null) NumberFormatId = pf.NumberFormatId.Value;
            if (pf.Outline != null) Outline = pf.Outline.Value;
            if (pf.SubtotalTop != null) SubtotalTop = pf.SubtotalTop.Value;
            if (pf.DragToRow != null) DragToRow = pf.DragToRow.Value;
            if (pf.DragToColumn != null) DragToColumn = pf.DragToColumn.Value;
            if (pf.MultipleItemSelectionAllowed != null)
                MultipleItemSelectionAllowed = pf.MultipleItemSelectionAllowed.Value;
            if (pf.DragToPage != null) DragToPage = pf.DragToPage.Value;
            if (pf.DragToData != null) DragToData = pf.DragToData.Value;
            if (pf.DragOff != null) DragOff = pf.DragOff.Value;
            if (pf.ShowAll != null) ShowAll = pf.ShowAll.Value;
            if (pf.InsertBlankRow != null) InsertBlankRow = pf.InsertBlankRow.Value;
            if (pf.ServerField != null) ServerField = pf.ServerField.Value;
            if (pf.InsertPageBreak != null) InsertPageBreak = pf.InsertPageBreak.Value;
            if (pf.AutoShow != null) AutoShow = pf.AutoShow.Value;
            if (pf.TopAutoShow != null) TopAutoShow = pf.TopAutoShow.Value;
            if (pf.HideNewItems != null) HideNewItems = pf.HideNewItems.Value;
            if (pf.MeasureFilter != null) MeasureFilter = pf.MeasureFilter.Value;
            if (pf.IncludeNewItemsInFilter != null) IncludeNewItemsInFilter = pf.IncludeNewItemsInFilter.Value;
            if (pf.ItemPageCount != null) ItemPageCount = pf.ItemPageCount.Value;
            if (pf.SortType != null) SortType = pf.SortType.Value;
            if (pf.DataSourceSort != null) DataSourceSort = pf.DataSourceSort.Value;
            if (pf.NonAutoSortDefault != null) NonAutoSortDefault = pf.NonAutoSortDefault.Value;
            if (pf.RankBy != null) RankBy = pf.RankBy.Value;
            if (pf.DefaultSubtotal != null) DefaultSubtotal = pf.DefaultSubtotal.Value;
            if (pf.SumSubtotal != null) SumSubtotal = pf.SumSubtotal.Value;
            if (pf.CountASubtotal != null) CountASubtotal = pf.CountASubtotal.Value;
            if (pf.AverageSubTotal != null) AverageSubTotal = pf.AverageSubTotal.Value;
            if (pf.MaxSubtotal != null) MaxSubtotal = pf.MaxSubtotal.Value;
            if (pf.MinSubtotal != null) MinSubtotal = pf.MinSubtotal.Value;
            if (pf.ApplyProductInSubtotal != null) ApplyProductInSubtotal = pf.ApplyProductInSubtotal.Value;
            if (pf.CountSubtotal != null) CountSubtotal = pf.CountSubtotal.Value;
            if (pf.ApplyStandardDeviationInSubtotal != null)
                ApplyStandardDeviationInSubtotal = pf.ApplyStandardDeviationInSubtotal.Value;
            if (pf.ApplyStandardDeviationPInSubtotal != null)
                ApplyStandardDeviationPInSubtotal = pf.ApplyStandardDeviationPInSubtotal.Value;
            if (pf.ApplyVarianceInSubtotal != null) ApplyVarianceInSubtotal = pf.ApplyVarianceInSubtotal.Value;
            if (pf.ApplyVariancePInSubtotal != null) ApplyVariancePInSubtotal = pf.ApplyVariancePInSubtotal.Value;
            if (pf.ShowPropCell != null) ShowPropCell = pf.ShowPropCell.Value;
            if (pf.ShowPropertyTooltip != null) ShowPropertyTooltip = pf.ShowPropertyTooltip.Value;
            if (pf.ShowPropAsCaption != null) ShowPropAsCaption = pf.ShowPropAsCaption.Value;
            if (pf.DefaultAttributeDrillState != null) DefaultAttributeDrillState = pf.DefaultAttributeDrillState.Value;

            SLItem it;
            using (var oxr = OpenXmlReader.Create(pf))
            {
                while (oxr.Read())
                    if (oxr.ElementType == typeof(Item))
                    {
                        it = new SLItem();
                        it.FromItem((Item) oxr.LoadCurrentElement());
                        Items.Add(it);
                    }
                    else if (oxr.ElementType == typeof(AutoSortScope))
                    {
                        AutoSortScope.FromAutoSortScope((AutoSortScope) oxr.LoadCurrentElement());
                        HasAutoSortScope = true;
                    }
            }
        }

        internal PivotField ToPivotField()
        {
            var pf = new PivotField();
            if ((Name != null) && (Name.Length > 0)) pf.Name = Name;
            if (Axis != null) pf.Axis = Axis.Value;
            if (DataField) pf.DataField = DataField;
            if ((SubtotalCaption != null) && (SubtotalCaption.Length > 0)) pf.SubtotalCaption = SubtotalCaption;
            if (ShowDropDowns != true) pf.ShowDropDowns = ShowDropDowns;
            if (HiddenLevel) pf.HiddenLevel = HiddenLevel;
            if ((UniqueMemberProperty != null) && (UniqueMemberProperty.Length > 0))
                pf.UniqueMemberProperty = UniqueMemberProperty;
            if (Compact != true) pf.Compact = Compact;
            if (AllDrilled) pf.AllDrilled = AllDrilled;
            if (NumberFormatId != null) pf.NumberFormatId = NumberFormatId.Value;
            if (Outline != true) pf.Outline = Outline;
            if (SubtotalTop != true) pf.SubtotalTop = SubtotalTop;
            if (DragToRow != true) pf.DragToRow = DragToRow;
            if (DragToColumn != true) pf.DragToColumn = DragToColumn;
            if (MultipleItemSelectionAllowed) pf.MultipleItemSelectionAllowed = MultipleItemSelectionAllowed;
            if (DragToPage != true) pf.DragToPage = DragToPage;
            if (DragToData != true) pf.DragToData = DragToData;
            if (DragOff != true) pf.DragOff = DragOff;
            if (ShowAll != true) pf.ShowAll = ShowAll;
            if (InsertBlankRow) pf.InsertBlankRow = InsertBlankRow;
            if (ServerField) pf.ServerField = ServerField;
            if (InsertPageBreak) pf.InsertPageBreak = InsertPageBreak;
            if (AutoShow) pf.AutoShow = AutoShow;
            if (TopAutoShow != true) pf.TopAutoShow = TopAutoShow;
            if (HideNewItems) pf.HideNewItems = HideNewItems;
            if (MeasureFilter) pf.MeasureFilter = MeasureFilter;
            if (IncludeNewItemsInFilter) pf.IncludeNewItemsInFilter = IncludeNewItemsInFilter;
            if (ItemPageCount != 10) pf.ItemPageCount = ItemPageCount;
            if (SortType != FieldSortValues.Manual) pf.SortType = SortType;
            if (DataSourceSort != null) pf.DataSourceSort = DataSourceSort.Value;
            if (NonAutoSortDefault) pf.NonAutoSortDefault = NonAutoSortDefault;
            if (RankBy != null) pf.RankBy = RankBy.Value;
            if (DefaultSubtotal != true) pf.DefaultSubtotal = DefaultSubtotal;
            if (SumSubtotal) pf.SumSubtotal = SumSubtotal;
            if (CountASubtotal) pf.CountASubtotal = CountASubtotal;
            if (AverageSubTotal) pf.AverageSubTotal = AverageSubTotal;
            if (MaxSubtotal) pf.MaxSubtotal = MaxSubtotal;
            if (MinSubtotal) pf.MinSubtotal = MinSubtotal;
            if (ApplyProductInSubtotal) pf.ApplyProductInSubtotal = ApplyProductInSubtotal;
            if (CountSubtotal) pf.CountSubtotal = CountSubtotal;
            if (ApplyStandardDeviationInSubtotal)
                pf.ApplyStandardDeviationInSubtotal = ApplyStandardDeviationInSubtotal;
            if (ApplyStandardDeviationPInSubtotal)
                pf.ApplyStandardDeviationPInSubtotal = ApplyStandardDeviationPInSubtotal;
            if (ApplyVarianceInSubtotal) pf.ApplyVarianceInSubtotal = ApplyVarianceInSubtotal;
            if (ApplyVariancePInSubtotal) pf.ApplyVariancePInSubtotal = ApplyVariancePInSubtotal;
            if (ShowPropCell) pf.ShowPropCell = ShowPropCell;
            if (ShowPropertyTooltip) pf.ShowPropertyTooltip = ShowPropertyTooltip;
            if (ShowPropAsCaption) pf.ShowPropAsCaption = ShowPropAsCaption;
            if (DefaultAttributeDrillState) pf.DefaultAttributeDrillState = DefaultAttributeDrillState;

            if (Items.Count > 0)
            {
                pf.Items = new Items();
                foreach (var it in Items)
                    pf.Items.Append(it.ToItem());
            }

            if (HasAutoSortScope)
                pf.AutoSortScope = AutoSortScope.ToAutoSortScope();

            return pf;
        }

        internal SLPivotField Clone()
        {
            var pf = new SLPivotField();
            pf.Name = Name;
            pf.Axis = Axis;
            pf.DataField = DataField;
            pf.SubtotalCaption = SubtotalCaption;
            pf.ShowDropDowns = ShowDropDowns;
            pf.HiddenLevel = HiddenLevel;
            pf.UniqueMemberProperty = UniqueMemberProperty;
            pf.Compact = Compact;
            pf.AllDrilled = AllDrilled;
            pf.NumberFormatId = NumberFormatId;
            pf.Outline = Outline;
            pf.SubtotalTop = SubtotalTop;
            pf.DragToRow = DragToRow;
            pf.DragToColumn = DragToColumn;
            pf.MultipleItemSelectionAllowed = MultipleItemSelectionAllowed;
            pf.DragToPage = DragToPage;
            pf.DragToData = DragToData;
            pf.DragOff = DragOff;
            pf.ShowAll = ShowAll;
            pf.InsertBlankRow = InsertBlankRow;
            pf.ServerField = ServerField;
            pf.InsertPageBreak = InsertPageBreak;
            pf.AutoShow = AutoShow;
            pf.TopAutoShow = TopAutoShow;
            pf.HideNewItems = HideNewItems;
            pf.MeasureFilter = MeasureFilter;
            pf.IncludeNewItemsInFilter = IncludeNewItemsInFilter;
            pf.ItemPageCount = ItemPageCount;
            pf.SortType = SortType;
            pf.DataSourceSort = DataSourceSort;
            pf.NonAutoSortDefault = NonAutoSortDefault;
            pf.RankBy = RankBy;
            pf.DefaultSubtotal = DefaultSubtotal;
            pf.SumSubtotal = SumSubtotal;
            pf.CountASubtotal = CountASubtotal;
            pf.AverageSubTotal = AverageSubTotal;
            pf.MaxSubtotal = MaxSubtotal;
            pf.MinSubtotal = MinSubtotal;
            pf.ApplyProductInSubtotal = ApplyProductInSubtotal;
            pf.CountSubtotal = CountSubtotal;
            pf.ApplyStandardDeviationInSubtotal = ApplyStandardDeviationInSubtotal;
            pf.ApplyStandardDeviationPInSubtotal = ApplyStandardDeviationPInSubtotal;
            pf.ApplyVarianceInSubtotal = ApplyVarianceInSubtotal;
            pf.ApplyVariancePInSubtotal = ApplyVariancePInSubtotal;
            pf.ShowPropCell = ShowPropCell;
            pf.ShowPropertyTooltip = ShowPropertyTooltip;
            pf.ShowPropAsCaption = ShowPropAsCaption;
            pf.DefaultAttributeDrillState = DefaultAttributeDrillState;

            pf.Items = new List<SLItem>();
            foreach (var it in Items)
                pf.Items.Add(it.Clone());

            pf.AutoSortScope = AutoSortScope.Clone();
            pf.HasAutoSortScope = HasAutoSortScope;

            return pf;
        }
    }
}