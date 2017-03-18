using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;
using Ups.Toolkit.SpreadsheetLight.Core.misc;
using Ups.Toolkit.SpreadsheetLight.Core.worksheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.pivottable
{
    /// <summary>
    ///     Data field function values.
    /// </summary>
    public enum SLDataFieldFunctionValues
    {
        /// <summary>
        ///     Average
        /// </summary>
        Average = 0,

        /// <summary>
        ///     Count non-empty cells
        /// </summary>
        Count,

        /// <summary>
        ///     Count numbers
        /// </summary>
        CountNumbers,

        /// <summary>
        ///     Maximum
        /// </summary>
        Maximum,

        /// <summary>
        ///     Minimum
        /// </summary>
        Minimum,

        /// <summary>
        ///     Product
        /// </summary>
        Product,

        /// <summary>
        ///     Standard deviation
        /// </summary>
        StandardDeviation,

        /// <summary>
        ///     Sum
        /// </summary>
        Sum,

        /// <summary>
        ///     Variance
        /// </summary>
        Variance
    }

    internal enum SLPivotFieldTypeValues
    {
        Filter = 0,
        Column,
        Row,
        Value,
        NotUsed
    }

    internal class SLPivotFieldType
    {
        internal SLPivotFieldType()
        {
            IsNumericIndex = true;
            FieldIndex = 0;
            FieldName = string.Empty;
            FieldType = SLPivotFieldTypeValues.NotUsed;
        }

        // determines whether to use FieldIndex or FieldName
        internal bool IsNumericIndex { get; set; }
        internal int FieldIndex { get; set; }
        internal string FieldName { get; set; }
        internal SLPivotFieldTypeValues FieldType { get; set; }
    }

    public class SLPivotTable
    {
        internal SLCellPointRange DataRange;

        /// <summary>
        ///     If true, then SheetTableName is a table name. Otherwise it's a worksheet name.
        /// </summary>
        internal bool IsDataSourceTable;

        internal bool IsNewPivotTable;
        internal string SheetTableName;

        internal SLPivotTable()
        {
            SetAllNull();
        }

        //CT_pivotTableDefinition
        //DocumentFormat.OpenXml.Spreadsheet.PivotTableDefinition

        //From Open XML specs: When encountering sheet boundaries, the PivotTable is truncated rather than wrapped, and as much as possible shall be shown.

        /*
         * <x:pivotCacheDefinition xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" refreshedBy="Vincent" refreshedDate="41315.775251967592" createdVersion="5" refreshedVersion="5" minRefreshableVersion="3" recordCount="5" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:cacheSource type="worksheet">
    <x:worksheetSource ref="A1:D6" sheet="Sheet1" />
  </x:cacheSource>


<x:pivotCacheDefinition xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" refreshedBy="Vincent" refreshedDate="41315.776955555557" createdVersion="5" refreshedVersion="5" minRefreshableVersion="3" recordCount="5" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <x:cacheSource type="worksheet">
    <x:worksheetSource name="Table1" />
  </x:cacheSource>
         * */

        // Pivot tables can have the same names as normal tables.

        internal bool IsValid
        {
            get
            {
                return (DataRange.StartRowIndex >= 1) && (DataRange.StartRowIndex <= SLConstants.RowLimit)
                       && ((DataRange.StartColumnIndex >= 1) & (DataRange.StartColumnIndex <= SLConstants.ColumnLimit))
                       && (DataRange.EndRowIndex >= 1) && (DataRange.EndRowIndex <= SLConstants.RowLimit)
                       && (DataRange.EndColumnIndex >= 1) && (DataRange.EndColumnIndex <= SLConstants.ColumnLimit);
            }
        }

        internal SLLocation Location { get; set; }
        internal List<SLPivotField> PivotFields { get; set; }
        internal List<int> RowFields { get; set; }
        internal List<SLRowItem> RowItems { get; set; }
        internal List<int> ColumnFields { get; set; }
        // ColumnItems use RowItem as children. Hey I'm not the one who designed this.
        internal List<SLRowItem> ColumnItems { get; set; }
        internal List<SLPageField> PageFields { get; set; }
        internal List<SLDataField> DataFields { get; set; }
        internal List<SLFormat> Formats { get; set; }
        internal List<SLConditionalFormat> ConditionalFormats { get; set; }
        internal List<SLChartFormat> ChartFormats { get; set; }
        internal List<SLPivotHierarchy> PivotHierarchies { get; set; }
        internal SLPivotTableStyle PivotTableStyle { get; set; }
        internal List<SLPivotFilter> PivotFilters { get; set; }
        internal List<int> RowHierarchiesUsage { get; set; }
        internal List<int> ColumnHierarchiesUsage { get; set; }

        // Oh my gamma rays the attributes are so *not* in accordance with the Open XML specs...
        //http://msdn.microsoft.com/en-us/library/ff532298%28v=office.12%29.aspx
        //http://msdn.microsoft.com/en-us/library/ff534910%28v=office.12%29.aspx

        //required attribute
        internal string Name { get; set; }
        //required attribute
        internal uint CacheId { get; set; }

        internal bool DataOnRows { get; set; }
        internal uint? DataPosition { get; set; }

        //required attribute
        internal string DataCaption { get; set; }

        internal string GrandTotalCaption { get; set; }
        internal string ErrorCaption { get; set; }
        internal bool ShowError { get; set; }
        internal string MissingCaption { get; set; }
        internal bool ShowMissing { get; set; }
        internal string PageStyle { get; set; }
        internal string PivotTableStyleName { get; set; }
        internal string VacatedStyle { get; set; }
        internal string Tag { get; set; }
        internal byte UpdatedVersion { get; set; }
        internal byte MinRefreshableVersion { get; set; }
        internal bool AsteriskTotals { get; set; }
        internal bool ShowItems { get; set; }
        internal bool EditData { get; set; }
        internal bool DisableFieldList { get; set; }
        internal bool ShowCalculatedMembers { get; set; }
        internal bool VisualTotals { get; set; }
        internal bool ShowMultipleLabel { get; set; }
        internal bool ShowDataDropDown { get; set; }
        internal bool ShowDrill { get; set; }
        internal bool PrintDrill { get; set; }
        internal bool ShowMemberPropertyTips { get; set; }
        internal bool ShowDataTips { get; set; }
        internal bool EnableWizard { get; set; }
        internal bool EnableDrill { get; set; }
        internal bool EnableFieldProperties { get; set; }
        internal bool PreserveFormatting { get; set; }
        internal bool UseAutoFormatting { get; set; }
        internal uint PageWrap { get; set; }
        internal bool PageOverThenDown { get; set; }
        internal bool SubtotalHiddenItems { get; set; }
        internal bool RowGrandTotals { get; set; }
        internal bool ColumnGrandTotals { get; set; }
        internal bool FieldPrintTitles { get; set; }
        internal bool ItemPrintTitles { get; set; }
        internal bool MergeItem { get; set; }
        internal bool ShowDropZones { get; set; }
        internal byte CreatedVersion { get; set; }
        internal uint Indent { get; set; }
        internal bool ShowEmptyRow { get; set; }
        internal bool ShowEmptyColumn { get; set; }
        internal bool ShowHeaders { get; set; }
        internal bool Compact { get; set; }
        internal bool Outline { get; set; }
        internal bool OutlineData { get; set; }
        internal bool CompactData { get; set; }
        internal bool Published { get; set; }
        internal bool GridDropZones { get; set; }
        internal bool StopImmersiveUi { get; set; }
        internal bool MultipleFieldFilters { get; set; }
        internal uint ChartFormat { get; set; }
        internal string RowHeaderCaption { get; set; }
        internal string ColumnHeaderCaption { get; set; }
        internal bool FieldListSortAscending { get; set; }
        // what happened to mdxSubqueries? It's in the Open XML specs...
        // See http://msdn.microsoft.com/en-us/library/ff532298%28v=office.12%29.aspx
        // "Office does not use the mdxSubqueries attribute". Really? Hurrumph.
        internal bool CustomListSort { get; set; }

        // we store the instructions for which field to be for which type.
        // Then when we actually render the pivot table at the point of inserting
        // into the worksheet, we'll use this instruction set to get and juggle the
        // data.
        // We do it this way so that if any other worksheet operations are done,
        // say insert rows or setting different cell values that happen to be in the
        // jurisdiction of the pivot table range, we don't have to rejuggle our pivot
        // table data because we haven't done anything yet!
        internal List<SLPivotFieldType> FieldSettingInstructions { get; set; }

        private void SetAllNull()
        {
            IsNewPivotTable = true;
            DataRange = new SLCellPointRange(-1, -1, -1, -1);
            IsDataSourceTable = false;
            SheetTableName = string.Empty;

            // Excel 2013 uses 5 for attributes createdVersion and updatedVersion
            // and 3 for attribute minRefreshableVersion.
            // I don't know what earlier versions of Excel use because I've uninstalled
            // the earlier versions of Excel 2007 and 2010...
            // These three attributes are application dependent, which means technically they're
            // dependent on SpreadsheetLight.
            // For the sake of simplicity, I'm going to set all three attributes to the value of 3,
            // so Excel (whatever version) can handle it without being insufferable.

            Name = "";
            CacheId = 0;
            DataOnRows = false;
            DataPosition = null;

            AutoFormatId = null;
            ApplyNumberFormats = null;
            ApplyBorderFormats = null;
            ApplyFontFormats = null;
            ApplyPatternFormats = null;
            ApplyAlignmentFormats = null;
            ApplyWidthHeightFormats = null;

            DataCaption = "";
            GrandTotalCaption = "";
            ErrorCaption = "";
            ShowError = false;
            MissingCaption = "";
            ShowMissing = true;
            PageStyle = "";
            PivotTableStyleName = "";
            VacatedStyle = "";
            Tag = "";
            UpdatedVersion = 3; // supposed to default 0. See above.
            MinRefreshableVersion = 3; // supposed to default 0. See above.
            AsteriskTotals = false;
            ShowItems = true;
            EditData = false;
            DisableFieldList = false;
            ShowCalculatedMembers = true;
            VisualTotals = true;
            ShowMultipleLabel = true;
            ShowDataDropDown = true;
            ShowDrill = true;
            PrintDrill = false;
            ShowMemberPropertyTips = true;
            ShowDataTips = true;
            EnableWizard = true;
            EnableDrill = true;
            EnableFieldProperties = true;
            PreserveFormatting = true;
            UseAutoFormatting = false;
            PageWrap = 0;
            PageOverThenDown = false;
            SubtotalHiddenItems = false;
            RowGrandTotals = true;
            ColumnGrandTotals = true;
            FieldPrintTitles = false;
            ItemPrintTitles = false;
            MergeItem = false;
            ShowDropZones = true;
            CreatedVersion = 3; // supposed to default 0. See above.
            Indent = 1;
            ShowEmptyRow = false;
            ShowEmptyColumn = false;
            ShowHeaders = true;
            Compact = true;
            Outline = false;
            OutlineData = false;
            CompactData = true;
            Published = false;
            GridDropZones = false;
            StopImmersiveUi = true;
            MultipleFieldFilters = true;
            ChartFormat = 0;
            RowHeaderCaption = "";
            ColumnHeaderCaption = "";
            FieldListSortAscending = false;
            CustomListSort = true;
        }

        public void SetFilterField(int FieldIndex)
        {
            var pft = new SLPivotFieldType();
            pft.IsNumericIndex = true;
            pft.FieldIndex = FieldIndex;
            pft.FieldType = SLPivotFieldTypeValues.Filter;
            FieldSettingInstructions.Add(pft);
        }

        public void SetFilterField(string FieldName)
        {
            var pft = new SLPivotFieldType();
            pft.IsNumericIndex = false;
            pft.FieldName = FieldName;
            pft.FieldType = SLPivotFieldTypeValues.Filter;
            FieldSettingInstructions.Add(pft);
        }

        public void SetColumnField(int FieldIndex)
        {
            var pft = new SLPivotFieldType();
            pft.IsNumericIndex = true;
            pft.FieldIndex = FieldIndex;
            pft.FieldType = SLPivotFieldTypeValues.Column;
            FieldSettingInstructions.Add(pft);
        }

        public void SetColumnField(string FieldName)
        {
            var pft = new SLPivotFieldType();
            pft.IsNumericIndex = false;
            pft.FieldName = FieldName;
            pft.FieldType = SLPivotFieldTypeValues.Column;
            FieldSettingInstructions.Add(pft);
        }

        public void SetRowField(int FieldIndex)
        {
            var pft = new SLPivotFieldType();
            pft.IsNumericIndex = true;
            pft.FieldIndex = FieldIndex;
            pft.FieldType = SLPivotFieldTypeValues.Row;
            FieldSettingInstructions.Add(pft);
        }

        public void SetRowField(string FieldName)
        {
            var pft = new SLPivotFieldType();
            pft.IsNumericIndex = false;
            pft.FieldName = FieldName;
            pft.FieldType = SLPivotFieldTypeValues.Row;
            FieldSettingInstructions.Add(pft);
        }

        public void SetValueField(int FieldIndex)
        {
            var pft = new SLPivotFieldType();
            pft.IsNumericIndex = true;
            pft.FieldIndex = FieldIndex;
            pft.FieldType = SLPivotFieldTypeValues.Value;
            FieldSettingInstructions.Add(pft);
        }

        public void SetValueField(string FieldName)
        {
            var pft = new SLPivotFieldType();
            pft.IsNumericIndex = false;
            pft.FieldName = FieldName;
            pft.FieldType = SLPivotFieldTypeValues.Value;
            FieldSettingInstructions.Add(pft);
        }

        internal PivotTableDefinition ToPivotTableDefinition()
        {
            var ptd = new PivotTableDefinition();

            ptd.Name = Name;
            ptd.CacheId = CacheId;
            if (DataOnRows) ptd.DataOnRows = DataOnRows;
            if (DataPosition != null) ptd.DataPosition = DataPosition.Value;

            if (AutoFormatId != null) ptd.AutoFormatId = AutoFormatId.Value;
            if (ApplyNumberFormats != null) ptd.ApplyNumberFormats = ApplyNumberFormats.Value;
            if (ApplyBorderFormats != null) ptd.ApplyBorderFormats = ApplyBorderFormats.Value;
            if (ApplyFontFormats != null) ptd.ApplyFontFormats = ApplyFontFormats.Value;
            if (ApplyPatternFormats != null) ptd.ApplyPatternFormats = ApplyPatternFormats.Value;
            if (ApplyAlignmentFormats != null) ptd.ApplyAlignmentFormats = ApplyAlignmentFormats.Value;
            if (ApplyWidthHeightFormats != null) ptd.ApplyWidthHeightFormats = ApplyWidthHeightFormats.Value;

            if ((DataCaption != null) && (DataCaption.Length > 0)) ptd.DataCaption = DataCaption;
            if ((GrandTotalCaption != null) && (GrandTotalCaption.Length > 0))
                ptd.GrandTotalCaption = GrandTotalCaption;
            if ((ErrorCaption != null) && (ErrorCaption.Length > 0)) ptd.ErrorCaption = ErrorCaption;
            if (ShowError) ptd.ShowError = ShowError;
            if ((MissingCaption != null) && (MissingCaption.Length > 0)) ptd.MissingCaption = MissingCaption;
            if (ShowMissing != true) ptd.ShowMissing = ShowMissing;
            if ((PageStyle != null) && (PageStyle.Length > 0)) ptd.PageStyle = PageStyle;
            if ((PivotTableStyleName != null) && (PivotTableStyleName.Length > 0))
                ptd.PivotTableStyleName = PivotTableStyleName;
            if ((VacatedStyle != null) && (VacatedStyle.Length > 0)) ptd.VacatedStyle = VacatedStyle;
            if ((Tag != null) && (Tag.Length > 0)) ptd.Tag = Tag;
            if (UpdatedVersion != 0) ptd.UpdatedVersion = UpdatedVersion;
            if (MinRefreshableVersion != 0) ptd.MinRefreshableVersion = MinRefreshableVersion;
            if (AsteriskTotals) ptd.AsteriskTotals = AsteriskTotals;
            if (ShowItems != true) ptd.ShowItems = ShowItems;
            if (EditData) ptd.EditData = EditData;
            if (DisableFieldList) ptd.DisableFieldList = DisableFieldList;
            if (ShowCalculatedMembers != true) ptd.ShowCalculatedMembers = ShowCalculatedMembers;
            if (VisualTotals != true) ptd.VisualTotals = VisualTotals;
            if (ShowMultipleLabel != true) ptd.ShowMultipleLabel = ShowMultipleLabel;
            if (ShowDataDropDown != true) ptd.ShowDataDropDown = ShowDataDropDown;
            if (ShowDrill != true) ptd.ShowDrill = ShowDrill;
            if (PrintDrill) ptd.PrintDrill = PrintDrill;
            if (ShowMemberPropertyTips != true) ptd.ShowMemberPropertyTips = ShowMemberPropertyTips;
            if (ShowDataTips != true) ptd.ShowDataTips = ShowDataTips;
            if (EnableWizard != true) ptd.EnableWizard = EnableWizard;
            if (EnableDrill != true) ptd.EnableDrill = EnableDrill;
            if (EnableFieldProperties != true) ptd.EnableFieldProperties = EnableFieldProperties;
            if (PreserveFormatting != true) ptd.PreserveFormatting = PreserveFormatting;
            if (UseAutoFormatting) ptd.UseAutoFormatting = UseAutoFormatting;
            if (PageWrap != 0) ptd.PageWrap = PageWrap;
            if (PageOverThenDown) ptd.PageOverThenDown = PageOverThenDown;
            if (SubtotalHiddenItems) ptd.SubtotalHiddenItems = SubtotalHiddenItems;
            if (RowGrandTotals != true) ptd.RowGrandTotals = RowGrandTotals;
            if (ColumnGrandTotals != true) ptd.ColumnGrandTotals = ColumnGrandTotals;
            if (FieldPrintTitles) ptd.FieldPrintTitles = FieldPrintTitles;
            if (ItemPrintTitles) ptd.ItemPrintTitles = ItemPrintTitles;
            if (MergeItem) ptd.MergeItem = MergeItem;
            if (ShowDropZones != true) ptd.ShowDropZones = ShowDropZones;
            if (CreatedVersion != 0) ptd.CreatedVersion = CreatedVersion;
            if (Indent != 1) ptd.Indent = Indent;
            if (ShowEmptyRow) ptd.ShowEmptyRow = ShowEmptyRow;
            if (ShowEmptyColumn) ptd.ShowEmptyColumn = ShowEmptyColumn;
            if (ShowHeaders != true) ptd.ShowHeaders = ShowHeaders;
            if (Compact != true) ptd.Compact = Compact;
            if (Outline) ptd.Outline = Outline;
            if (OutlineData) ptd.OutlineData = OutlineData;
            if (CompactData != true) ptd.CompactData = CompactData;
            if (Published) ptd.Published = Published;
            if (GridDropZones) ptd.GridDropZones = GridDropZones;
            if (StopImmersiveUi != true) ptd.StopImmersiveUi = StopImmersiveUi;
            if (MultipleFieldFilters != true) ptd.MultipleFieldFilters = MultipleFieldFilters;
            if (ChartFormat != 0) ptd.ChartFormat = ChartFormat;
            if ((RowHeaderCaption != null) && (RowHeaderCaption.Length > 0)) ptd.RowHeaderCaption = RowHeaderCaption;
            if ((ColumnHeaderCaption != null) && (ColumnHeaderCaption.Length > 0))
                ptd.ColumnHeaderCaption = ColumnHeaderCaption;
            if (FieldListSortAscending) ptd.FieldListSortAscending = FieldListSortAscending;
            if (CustomListSort != true) ptd.CustomListSort = CustomListSort;

            ptd.Location = Location.ToLocation();

            if (PivotFields.Count > 0)
            {
                ptd.PivotFields = new PivotFields {Count = (uint) PivotFields.Count};
                foreach (var pf in PivotFields)
                    ptd.PivotFields.Append(pf.ToPivotField());
            }

            if (RowFields.Count > 0)
            {
                ptd.RowFields = new RowFields {Count = (uint) RowFields.Count};
                foreach (var i in RowFields)
                    ptd.RowFields.Append(new Field {Index = i});
            }

            if (RowItems.Count > 0)
            {
                ptd.RowItems = new RowItems {Count = (uint) RowItems.Count};
                foreach (var ri in RowItems)
                    ptd.RowItems.Append(ri.ToRowItem());
            }

            if (ColumnFields.Count > 0)
            {
                ptd.ColumnFields = new ColumnFields {Count = (uint) ColumnFields.Count};
                foreach (var i in ColumnFields)
                    ptd.ColumnFields.Append(new Field {Index = i});
            }

            if (ColumnItems.Count > 0)
            {
                ptd.ColumnItems = new ColumnItems {Count = (uint) ColumnItems.Count};
                foreach (var ri in ColumnItems)
                    ptd.ColumnItems.Append(ri.ToRowItem());
            }

            if (PageFields.Count > 0)
            {
                ptd.PageFields = new PageFields {Count = (uint) PageFields.Count};
                foreach (var pf in PageFields)
                    ptd.PageFields.Append(pf.ToPageField());
            }

            if (DataFields.Count > 0)
            {
                ptd.DataFields = new DataFields {Count = (uint) DataFields.Count};
                foreach (var df in DataFields)
                    ptd.DataFields.Append(df.ToDataField());
            }

            if (Formats.Count > 0)
            {
                ptd.Formats = new Formats {Count = (uint) Formats.Count};
                foreach (var f in Formats)
                    ptd.Formats.Append(f.ToFormat());
            }

            if (ConditionalFormats.Count > 0)
            {
                ptd.ConditionalFormats = new ConditionalFormats {Count = (uint) ConditionalFormats.Count};
                foreach (var cf in ConditionalFormats)
                    ptd.ConditionalFormats.Append(cf.ToConditionalFormat());
            }

            if (ChartFormats.Count > 0)
            {
                ptd.ChartFormats = new ChartFormats {Count = (uint) ChartFormats.Count};
                foreach (var cf in ChartFormats)
                    ptd.ChartFormats.Append(cf.ToChartFormat());
            }

            if (PivotHierarchies.Count > 0)
            {
                ptd.PivotHierarchies = new PivotHierarchies {Count = (uint) PivotHierarchies.Count};
                foreach (var ph in PivotHierarchies)
                    ptd.PivotHierarchies.Append(ph.ToPivotHierarchy());
            }

            ptd.PivotTableStyle = PivotTableStyle.ToPivotTableStyle();

            if (PivotFilters.Count > 0)
            {
                ptd.PivotFilters = new PivotFilters {Count = (uint) PivotFilters.Count};
                foreach (var pf in PivotFilters)
                    ptd.PivotFilters.Append(pf.ToPivotFilter());
            }

            if (RowHierarchiesUsage.Count > 0)
            {
                ptd.RowHierarchiesUsage = new RowHierarchiesUsage {Count = (uint) RowHierarchiesUsage.Count};
                foreach (var i in RowHierarchiesUsage)
                    ptd.RowHierarchiesUsage.Append(new RowHierarchyUsage {Value = i});
            }

            if (ColumnHierarchiesUsage.Count > 0)
            {
                ptd.ColumnHierarchiesUsage = new ColumnHierarchiesUsage {Count = (uint) ColumnHierarchiesUsage.Count};
                foreach (var i in ColumnHierarchiesUsage)
                    ptd.ColumnHierarchiesUsage.Append(new ColumnHierarchyUsage {Value = i});
            }

            return ptd;
        }

        #region AG_AutoFormat

        internal uint? AutoFormatId { get; set; }
        internal bool? ApplyNumberFormats { get; set; }
        internal bool? ApplyBorderFormats { get; set; }
        internal bool? ApplyFontFormats { get; set; }
        internal bool? ApplyPatternFormats { get; set; }
        internal bool? ApplyAlignmentFormats { get; set; }
        internal bool? ApplyWidthHeightFormats { get; set; }

        #endregion
    }
}