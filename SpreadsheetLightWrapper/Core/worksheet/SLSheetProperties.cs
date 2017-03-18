using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLightWrapper.Core.style;
using Color = System.Drawing.Color;

namespace SpreadsheetLightWrapper.Core.worksheet
{
    internal class SLSheetProperties
    {
        internal SLColor clrTabColor;

        internal bool HasTabColor;
        internal List<Color> listIndexedColors;

        internal List<Color> listThemeColors;

        internal SLSheetProperties(List<Color> ThemeColors, List<Color> IndexedColors)
        {
            int i;
            listThemeColors = new List<Color>();
            for (i = 0; i < ThemeColors.Count; ++i)
                listThemeColors.Add(ThemeColors[i]);

            listIndexedColors = new List<Color>();
            for (i = 0; i < IndexedColors.Count; ++i)
                listIndexedColors.Add(IndexedColors[i]);

            SetAllNull();
        }

        internal bool HasSheetProperties
        {
            get
            {
                return HasTabColor || ApplyStyles || !SummaryBelow
                       || !SummaryRight || !ShowOutlineSymbols
                       || !AutoPageBreaks || FitToPage
                       || SyncHorizontal || SyncVertical || (SyncReference.Length > 0)
                       || TransitionEvaluation || TransitionEntry
                       || !Published || (CodeName.Length > 0) || FilterMode
                       || !EnableFormatConditionsCalculation;
            }
        }

        internal bool HasChartSheetProperties
        {
            get { return HasTabColor || !Published || (CodeName.Length > 0); }
        }

        internal Color TabColor
        {
            get { return clrTabColor.Color; }
            set
            {
                clrTabColor.Color = value;
                HasTabColor = clrTabColor.Color.IsEmpty ? false : true;
            }
        }

        internal bool ApplyStyles { get; set; }
        internal bool SummaryBelow { get; set; }
        internal bool SummaryRight { get; set; }
        internal bool ShowOutlineSymbols { get; set; }

        internal bool AutoPageBreaks { get; set; }
        internal bool FitToPage { get; set; }

        internal bool SyncHorizontal { get; set; }
        internal bool SyncVertical { get; set; }
        internal string SyncReference { get; set; }
        internal bool TransitionEvaluation { get; set; }
        internal bool TransitionEntry { get; set; }
        internal bool Published { get; set; }
        internal string CodeName { get; set; }
        internal bool FilterMode { get; set; }
        internal bool EnableFormatConditionsCalculation { get; set; }

        private void SetAllNull()
        {
            clrTabColor = new SLColor(listThemeColors, listIndexedColors);
            HasTabColor = false;

            ApplyStyles = false;
            SummaryBelow = false;
            SummaryRight = false;
            ShowOutlineSymbols = true;

            AutoPageBreaks = true;
            FitToPage = false;

            SyncHorizontal = false;
            SyncVertical = false;
            SyncReference = string.Empty;
            TransitionEvaluation = false;
            TransitionEntry = false;
            Published = true;
            CodeName = string.Empty;
            FilterMode = false;
            EnableFormatConditionsCalculation = true;
        }

        internal void FromSheetProperties(SheetProperties sp)
        {
            SetAllNull();
            if (sp.TabColor != null)
                if ((sp.TabColor.Indexed != null) || (sp.TabColor.Theme != null) || (sp.TabColor.Rgb != null))
                {
                    clrTabColor.FromTabColor(sp.TabColor);
                    HasTabColor = clrTabColor.Color.IsEmpty ? false : true;
                }

            if (sp.OutlineProperties != null)
            {
                if (sp.OutlineProperties.ApplyStyles != null) ApplyStyles = sp.OutlineProperties.ApplyStyles.Value;
                if (sp.OutlineProperties.SummaryBelow != null) SummaryBelow = sp.OutlineProperties.SummaryBelow.Value;
                if (sp.OutlineProperties.SummaryRight != null) SummaryRight = sp.OutlineProperties.SummaryRight.Value;
                if (sp.OutlineProperties.ShowOutlineSymbols != null)
                    ShowOutlineSymbols = sp.OutlineProperties.ShowOutlineSymbols.Value;
            }

            if (sp.PageSetupProperties != null)
            {
                if (sp.PageSetupProperties.AutoPageBreaks != null)
                    AutoPageBreaks = sp.PageSetupProperties.AutoPageBreaks.Value;
                if (sp.PageSetupProperties.FitToPage != null) FitToPage = sp.PageSetupProperties.FitToPage.Value;
            }

            if (sp.SyncHorizontal != null) SyncHorizontal = sp.SyncHorizontal.Value;
            if (sp.SyncVertical != null) SyncVertical = sp.SyncVertical.Value;
            if (sp.SyncReference != null) SyncReference = sp.SyncReference.Value;
            if (sp.TransitionEvaluation != null) TransitionEvaluation = sp.TransitionEvaluation.Value;
            if (sp.TransitionEntry != null) TransitionEntry = sp.TransitionEntry.Value;
            if (sp.Published != null) Published = sp.Published.Value;
            if (sp.CodeName != null) CodeName = sp.CodeName.Value;
            if (sp.FilterMode != null) FilterMode = sp.FilterMode.Value;
            if (sp.EnableFormatConditionsCalculation != null)
                EnableFormatConditionsCalculation = sp.EnableFormatConditionsCalculation.Value;
        }

        internal SheetProperties ToSheetProperties()
        {
            var sp = new SheetProperties();

            if (HasTabColor)
                sp.TabColor = clrTabColor.ToTabColor();

            if (ApplyStyles || !SummaryBelow || !SummaryRight || !ShowOutlineSymbols)
            {
                sp.OutlineProperties = new OutlineProperties();
                if (ApplyStyles) sp.OutlineProperties.ApplyStyles = ApplyStyles;
                if (!SummaryBelow) sp.OutlineProperties.SummaryBelow = SummaryBelow;
                if (!SummaryRight) sp.OutlineProperties.SummaryRight = SummaryRight;
                if (!ShowOutlineSymbols) sp.OutlineProperties.ShowOutlineSymbols = ShowOutlineSymbols;
            }

            if (!AutoPageBreaks || FitToPage)
            {
                sp.PageSetupProperties = new PageSetupProperties();
                if (!AutoPageBreaks) sp.PageSetupProperties.AutoPageBreaks = AutoPageBreaks;
                if (FitToPage) sp.PageSetupProperties.FitToPage = FitToPage;
            }

            if (SyncHorizontal) sp.SyncHorizontal = SyncHorizontal;
            if (SyncVertical) sp.SyncVertical = SyncVertical;
            if (SyncReference.Length > 0) sp.SyncReference = SyncReference;
            if (TransitionEvaluation) sp.TransitionEvaluation = TransitionEvaluation;
            if (TransitionEntry) sp.TransitionEntry = TransitionEntry;
            if (!Published) sp.Published = Published;
            if (CodeName.Length > 0) sp.CodeName = CodeName;
            if (FilterMode) sp.FilterMode = FilterMode;
            if (!EnableFormatConditionsCalculation)
                sp.EnableFormatConditionsCalculation = EnableFormatConditionsCalculation;

            return sp;
        }

        internal void FromChartSheetProperties(ChartSheetProperties sp)
        {
            SetAllNull();
            if (sp.TabColor != null)
                if ((sp.TabColor.Indexed != null) || (sp.TabColor.Theme != null) || (sp.TabColor.Rgb != null))
                {
                    clrTabColor.FromTabColor(sp.TabColor);
                    HasTabColor = clrTabColor.Color.IsEmpty ? false : true;
                }

            if (sp.Published != null) Published = sp.Published.Value;
            if (sp.CodeName != null) CodeName = sp.CodeName.Value;
        }

        internal ChartSheetProperties ToChartSheetProperties()
        {
            var csp = new ChartSheetProperties();

            if (HasTabColor)
                csp.TabColor = clrTabColor.ToTabColor();

            if (!Published) csp.Published = Published;
            if (CodeName.Length > 0) csp.CodeName = CodeName;

            return csp;
        }

        internal SLSheetProperties Clone()
        {
            var sp = new SLSheetProperties(listThemeColors, listIndexedColors);
            sp.clrTabColor = clrTabColor.Clone();
            sp.HasTabColor = HasTabColor;

            sp.ApplyStyles = ApplyStyles;
            sp.SummaryBelow = SummaryBelow;
            sp.SummaryRight = SummaryRight;
            sp.ShowOutlineSymbols = ShowOutlineSymbols;

            sp.AutoPageBreaks = AutoPageBreaks;
            sp.FitToPage = FitToPage;

            sp.SyncHorizontal = SyncHorizontal;
            sp.SyncVertical = SyncVertical;
            sp.SyncReference = SyncReference;
            sp.TransitionEvaluation = TransitionEvaluation;
            sp.TransitionEntry = TransitionEntry;
            sp.Published = Published;
            sp.CodeName = CodeName;
            sp.FilterMode = FilterMode;
            sp.EnableFormatConditionsCalculation = EnableFormatConditionsCalculation;

            return sp;
        }
    }
}