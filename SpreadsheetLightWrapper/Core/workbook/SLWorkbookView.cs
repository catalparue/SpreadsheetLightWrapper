using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.workbook
{
    internal class SLWorkbookView
    {
        internal SLWorkbookView()
        {
            SetAllNull();
        }

        internal VisibilityValues Visibility { get; set; }
        internal bool Minimized { get; set; }
        internal bool ShowHorizontalScroll { get; set; }
        internal bool ShowVerticalScroll { get; set; }
        internal bool ShowSheetTabs { get; set; }
        internal int? XWindow { get; set; }
        internal int? YWindow { get; set; }
        internal uint? WindowWidth { get; set; }
        internal uint? WindowHeight { get; set; }
        internal uint TabRatio { get; set; }
        internal uint FirstSheet { get; set; }
        internal uint ActiveTab { get; set; }
        internal bool AutoFilterDateGrouping { get; set; }

        private void SetAllNull()
        {
            Visibility = VisibilityValues.Visible;
            Minimized = false;
            ShowHorizontalScroll = true;
            ShowVerticalScroll = true;
            ShowSheetTabs = true;
            XWindow = null;
            YWindow = null;
            WindowWidth = null;
            WindowHeight = null;
            TabRatio = 600;
            FirstSheet = 0;
            ActiveTab = 0;
            AutoFilterDateGrouping = true;
        }

        internal void FromWorkbookView(WorkbookView wv)
        {
            SetAllNull();

            if (wv.Visibility != null) Visibility = wv.Visibility.Value;
            if (wv.Minimized != null) Minimized = wv.Minimized.Value;
            if (wv.ShowHorizontalScroll != null) ShowHorizontalScroll = wv.ShowHorizontalScroll.Value;
            if (wv.ShowVerticalScroll != null) ShowVerticalScroll = wv.ShowVerticalScroll.Value;
            if (wv.ShowSheetTabs != null) ShowSheetTabs = wv.ShowSheetTabs.Value;
            if (wv.XWindow != null) XWindow = wv.XWindow.Value;
            if (wv.YWindow != null) YWindow = wv.YWindow.Value;
            if (wv.WindowWidth != null) WindowWidth = wv.WindowWidth.Value;
            if (wv.WindowHeight != null) WindowHeight = wv.WindowHeight.Value;
            if (wv.TabRatio != null) TabRatio = wv.TabRatio.Value;
            if (wv.FirstSheet != null) FirstSheet = wv.FirstSheet.Value;
            if (wv.ActiveTab != null) ActiveTab = wv.ActiveTab.Value;
            if (wv.AutoFilterDateGrouping != null) AutoFilterDateGrouping = wv.AutoFilterDateGrouping.Value;
        }

        internal WorkbookView ToWorkbookView()
        {
            var wv = new WorkbookView();
            if (Visibility != VisibilityValues.Visible) wv.Visibility = Visibility;
            if (Minimized) wv.Minimized = Minimized;
            if (!ShowHorizontalScroll) wv.ShowHorizontalScroll = ShowHorizontalScroll;
            if (!ShowVerticalScroll) wv.ShowVerticalScroll = ShowVerticalScroll;
            if (!ShowSheetTabs) wv.ShowSheetTabs = ShowSheetTabs;
            if (XWindow != null) wv.XWindow = XWindow.Value;
            if (YWindow != null) wv.YWindow = YWindow.Value;
            if (WindowWidth != null) wv.WindowWidth = WindowWidth.Value;
            if (WindowHeight != null) wv.WindowHeight = WindowHeight.Value;
            if (TabRatio != 600) wv.TabRatio = TabRatio;
            if (FirstSheet != 0) wv.FirstSheet = FirstSheet;
            if (ActiveTab != 0) wv.ActiveTab = ActiveTab;
            if (!AutoFilterDateGrouping) wv.AutoFilterDateGrouping = AutoFilterDateGrouping;

            return wv;
        }

        internal SLWorkbookView Clone()
        {
            var wv = new SLWorkbookView();
            wv.Visibility = Visibility;
            wv.Minimized = Minimized;
            wv.ShowHorizontalScroll = ShowHorizontalScroll;
            wv.ShowVerticalScroll = ShowVerticalScroll;
            wv.ShowSheetTabs = ShowSheetTabs;
            wv.XWindow = XWindow;
            wv.YWindow = YWindow;
            wv.WindowWidth = WindowWidth;
            wv.WindowHeight = WindowHeight;
            wv.TabRatio = TabRatio;
            wv.FirstSheet = FirstSheet;
            wv.ActiveTab = ActiveTab;
            wv.AutoFilterDateGrouping = AutoFilterDateGrouping;

            return wv;
        }
    }
}