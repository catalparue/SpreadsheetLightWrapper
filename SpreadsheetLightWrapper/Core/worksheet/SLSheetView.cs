using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.worksheet
{
    internal class SLSheetView
    {
        internal uint iZoomScale;

        internal uint iZoomScaleNormal;

        internal uint iZoomScalePageLayoutView;

        internal uint iZoomScaleSheetLayoutView;

        internal SLSheetView()
        {
            SetAllNull();
        }

        // Part of the set of properties is also in SLPageSettings.
        // Short answer is that it's easier for the developer to access all page-like
        // properties in one class. That's why this class isn't exposed.
        // So remember to sync whenever relevant.

        internal bool HasPane
        {
            get
            {
                return (Pane.HorizontalSplit != 0) || (Pane.VerticalSplit != 0)
                       || (Pane.TopLeftCell != null) || (Pane.ActivePane != PaneValues.TopLeft)
                       || (Pane.State != PaneStateValues.Split);
            }
        }

        internal SLPane Pane { get; set; }

        internal List<SLSelection> Selections { get; set; }
        internal List<PivotSelection> PivotSelections { get; set; }

        internal bool WindowProtection { get; set; }

        // it appears that this doesn't change the showing/hiding of formula bar in Excel.
        // Excel has an option that does this, but is outside of the control of the Open XML
        // document. So it doesn't matter if you're using Open XML SDK or just rendering
        // XML tags. You're not going to control it.
        // Why? I don't know. Ask the Microsoft developer responsible.
        // It appears that if you change that option (within Excel), *all* Excel spreadsheets either
        // show or hide the formula bar for all worksheets.
        // Why this option is even here is beyond me...
        internal bool ShowFormulas { get; set; }

        internal bool ShowGridLines { get; set; }
        internal bool ShowRowColHeaders { get; set; }
        internal bool ShowZeros { get; set; }
        internal bool RightToLeft { get; set; }
        internal bool TabSelected { get; set; }
        internal bool ShowRuler { get; set; }
        internal bool ShowOutlineSymbols { get; set; }
        internal bool DefaultGridColor { get; set; }
        internal bool ShowWhiteSpace { get; set; }
        internal SheetViewValues View { get; set; }
        internal string TopLeftCell { get; set; }
        internal uint ColorId { get; set; }

        internal uint ZoomScale
        {
            get { return iZoomScale; }
            set
            {
                iZoomScale = value;
                if (iZoomScale < 10) iZoomScale = 10;
                if (iZoomScale > 400) iZoomScale = 400;
            }
        }

        internal uint ZoomScaleNormal
        {
            get { return iZoomScaleNormal; }
            set
            {
                iZoomScaleNormal = value;
                if (iZoomScaleNormal < 10) iZoomScaleNormal = 10;
                if (iZoomScaleNormal > 400) iZoomScaleNormal = 400;
            }
        }

        internal uint ZoomScaleSheetLayoutView
        {
            get { return iZoomScaleSheetLayoutView; }
            set
            {
                iZoomScaleSheetLayoutView = value;
                if (iZoomScaleSheetLayoutView < 10) iZoomScaleSheetLayoutView = 10;
                if (iZoomScaleSheetLayoutView > 400) iZoomScaleSheetLayoutView = 400;
            }
        }

        internal uint ZoomScalePageLayoutView
        {
            get { return iZoomScalePageLayoutView; }
            set
            {
                iZoomScalePageLayoutView = value;
                if (iZoomScalePageLayoutView < 10) iZoomScalePageLayoutView = 10;
                if (iZoomScalePageLayoutView > 400) iZoomScalePageLayoutView = 400;
            }
        }

        internal uint WorkbookViewId { get; set; }

        private void SetAllNull()
        {
            Pane = new SLPane();

            Selections = new List<SLSelection>();
            PivotSelections = new List<PivotSelection>();

            WindowProtection = false;
            ShowFormulas = false;
            ShowGridLines = true;
            ShowRowColHeaders = true;
            ShowZeros = true;
            RightToLeft = false;
            TabSelected = false;
            ShowRuler = true;
            ShowOutlineSymbols = true;
            DefaultGridColor = true;
            ShowWhiteSpace = true;
            View = SheetViewValues.Normal;
            TopLeftCell = string.Empty;
            ColorId = 64;
            iZoomScale = 100;
            iZoomScaleNormal = 0;
            iZoomScaleSheetLayoutView = 0;
            iZoomScalePageLayoutView = 0;
            WorkbookViewId = 0;
        }

        internal void FromSheetView(SheetView sv)
        {
            SetAllNull();

            if (sv.WindowProtection != null) WindowProtection = sv.WindowProtection.Value;
            if (sv.ShowFormulas != null) ShowFormulas = sv.ShowFormulas.Value;
            if (sv.ShowGridLines != null) ShowGridLines = sv.ShowGridLines.Value;
            if (sv.ShowRowColHeaders != null) ShowRowColHeaders = sv.ShowRowColHeaders.Value;
            if (sv.ShowZeros != null) ShowZeros = sv.ShowZeros.Value;
            if (sv.RightToLeft != null) RightToLeft = sv.RightToLeft.Value;
            if (sv.TabSelected != null) TabSelected = sv.TabSelected.Value;
            if (sv.ShowRuler != null) ShowRuler = sv.ShowRuler.Value;
            if (sv.ShowOutlineSymbols != null) ShowOutlineSymbols = sv.ShowOutlineSymbols.Value;
            if (sv.DefaultGridColor != null) DefaultGridColor = sv.DefaultGridColor.Value;
            if (sv.ShowWhiteSpace != null) ShowWhiteSpace = sv.ShowWhiteSpace.Value;
            if (sv.View != null) View = sv.View.Value;
            if (sv.TopLeftCell != null) TopLeftCell = sv.TopLeftCell.Value;
            if (sv.ColorId != null) ColorId = sv.ColorId.Value;
            if (sv.ZoomScale != null) ZoomScale = sv.ZoomScale.Value;
            if (sv.ZoomScaleNormal != null) ZoomScaleNormal = sv.ZoomScaleNormal.Value;
            if (sv.ZoomScaleSheetLayoutView != null) ZoomScaleSheetLayoutView = sv.ZoomScaleSheetLayoutView.Value;
            if (sv.ZoomScalePageLayoutView != null) ZoomScalePageLayoutView = sv.ZoomScalePageLayoutView.Value;

            // required attribute but we'll use 0 as the default in case something terrible happens.
            if (sv.WorkbookViewId != null) WorkbookViewId = sv.WorkbookViewId.Value;
            else WorkbookViewId = 0;

            using (var oxr = OpenXmlReader.Create(sv))
            {
                SLSelection sel;
                while (oxr.Read())
                    if (oxr.ElementType == typeof(Pane))
                    {
                        Pane = new SLPane();
                        Pane.FromPane((Pane) oxr.LoadCurrentElement());
                    }
                    else if (oxr.ElementType == typeof(Selection))
                    {
                        sel = new SLSelection();
                        sel.FromSelection((Selection) oxr.LoadCurrentElement());
                        Selections.Add(sel);
                    }
                    else if (oxr.ElementType == typeof(PivotSelection))
                    {
                        PivotSelections.Add((PivotSelection) oxr.LoadCurrentElement().CloneNode(true));
                    }
            }
        }

        internal SheetView ToSheetView()
        {
            var sv = new SheetView();
            if (WindowProtection) sv.WindowProtection = WindowProtection;
            if (ShowFormulas) sv.ShowFormulas = ShowFormulas;
            if (ShowGridLines != true) sv.ShowGridLines = ShowGridLines;
            if (ShowRowColHeaders != true) sv.ShowRowColHeaders = ShowRowColHeaders;
            if (ShowZeros != true) sv.ShowZeros = ShowZeros;
            if (RightToLeft) sv.RightToLeft = RightToLeft;
            if (TabSelected) sv.TabSelected = TabSelected;
            if (ShowRuler != true) sv.ShowRuler = ShowRuler;
            if (ShowOutlineSymbols != true) sv.ShowOutlineSymbols = ShowOutlineSymbols;
            if (DefaultGridColor != true) sv.DefaultGridColor = DefaultGridColor;
            if (ShowWhiteSpace != true) sv.ShowWhiteSpace = ShowWhiteSpace;
            if (View != SheetViewValues.Normal) sv.View = View;
            if ((TopLeftCell != null) && (TopLeftCell.Length > 0)) sv.TopLeftCell = TopLeftCell;
            if (ColorId != 64) sv.ColorId = ColorId;
            if (ZoomScale != 100) sv.ZoomScale = ZoomScale;
            if (ZoomScaleNormal != 0) sv.ZoomScaleNormal = ZoomScaleNormal;
            if (ZoomScaleSheetLayoutView != 0) sv.ZoomScaleSheetLayoutView = ZoomScaleSheetLayoutView;
            if (ZoomScalePageLayoutView != 0) sv.ZoomScalePageLayoutView = ZoomScalePageLayoutView;
            sv.WorkbookViewId = WorkbookViewId;

            if (HasPane)
                sv.Append(Pane.ToPane());

            foreach (var sel in Selections)
                sv.Append(sel.ToSelection());

            foreach (var psel in PivotSelections)
                sv.Append((PivotSelection) psel.CloneNode(true));

            return sv;
        }

        internal SLSheetView Clone()
        {
            var sv = new SLSheetView();
            sv.Pane = Pane.Clone();

            sv.Selections = new List<SLSelection>();
            foreach (var sel in Selections)
                sv.Selections.Add(sel.Clone());

            sv.PivotSelections = new List<PivotSelection>();
            foreach (var psel in PivotSelections)
                sv.PivotSelections.Add((PivotSelection) psel.CloneNode(true));

            sv.WindowProtection = WindowProtection;
            sv.ShowFormulas = ShowFormulas;
            sv.ShowGridLines = ShowGridLines;
            sv.ShowRowColHeaders = ShowRowColHeaders;
            sv.ShowZeros = ShowZeros;
            sv.RightToLeft = RightToLeft;
            sv.TabSelected = TabSelected;
            sv.ShowRuler = ShowRuler;
            sv.ShowOutlineSymbols = ShowOutlineSymbols;
            sv.DefaultGridColor = DefaultGridColor;
            sv.ShowWhiteSpace = ShowWhiteSpace;
            sv.View = View;
            sv.TopLeftCell = TopLeftCell;
            sv.ColorId = ColorId;
            sv.iZoomScale = iZoomScale;
            sv.iZoomScaleNormal = iZoomScaleNormal;
            sv.iZoomScaleSheetLayoutView = iZoomScaleSheetLayoutView;
            sv.iZoomScalePageLayoutView = iZoomScalePageLayoutView;
            sv.WorkbookViewId = WorkbookViewId;

            return sv;
        }

        internal static string GetSheetViewValuesAttribute(SheetViewValues svv)
        {
            var result = "normal";
            switch (svv)
            {
                case SheetViewValues.Normal:
                    result = "normal";
                    break;
                case SheetViewValues.PageBreakPreview:
                    result = "pageBreakPreview";
                    break;
                case SheetViewValues.PageLayout:
                    result = "pageLayout";
                    break;
            }

            return result;
        }
    }
}