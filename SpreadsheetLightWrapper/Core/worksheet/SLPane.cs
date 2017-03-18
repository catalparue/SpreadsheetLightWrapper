using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.worksheet
{
    internal class SLPane
    {
        internal SLPane()
        {
            SetAllNull();
        }

        internal double HorizontalSplit { get; set; }
        internal double VerticalSplit { get; set; }
        internal string TopLeftCell { get; set; }
        internal PaneValues ActivePane { get; set; }
        internal PaneStateValues State { get; set; }

        private void SetAllNull()
        {
            HorizontalSplit = 0;
            VerticalSplit = 0;
            TopLeftCell = null;
            ActivePane = PaneValues.TopLeft;
            State = PaneStateValues.Split;
        }

        internal void FromPane(Pane p)
        {
            SetAllNull();

            if (p.HorizontalSplit != null) HorizontalSplit = p.HorizontalSplit.Value;
            if (p.VerticalSplit != null) VerticalSplit = p.VerticalSplit.Value;
            if (p.TopLeftCell != null) TopLeftCell = p.TopLeftCell.Value;
            if (p.ActivePane != null) ActivePane = p.ActivePane.Value;
            if (p.State != null) State = p.State.Value;
        }

        internal Pane ToPane()
        {
            var p = new Pane();
            if (HorizontalSplit != 0) p.HorizontalSplit = HorizontalSplit;
            if (VerticalSplit != 0) p.VerticalSplit = VerticalSplit;
            if ((TopLeftCell != null) && (TopLeftCell.Length > 0)) p.TopLeftCell = TopLeftCell;
            if (ActivePane != PaneValues.TopLeft) p.ActivePane = ActivePane;
            if (State != PaneStateValues.Split) p.State = State;

            return p;
        }

        internal SLPane Clone()
        {
            var p = new SLPane();
            p.HorizontalSplit = HorizontalSplit;
            p.VerticalSplit = VerticalSplit;
            p.TopLeftCell = TopLeftCell;
            p.ActivePane = ActivePane;
            p.State = State;

            return p;
        }

        internal static string GetPaneValuesAttribute(PaneValues pv)
        {
            var result = "topLeft";
            switch (pv)
            {
                case PaneValues.BottomLeft:
                    result = "bottomLeft";
                    break;
                case PaneValues.BottomRight:
                    result = "bottomRight";
                    break;
                case PaneValues.TopLeft:
                    result = "topLeft";
                    break;
                case PaneValues.TopRight:
                    result = "topRight";
                    break;
            }

            return result;
        }

        internal static string GetPaneStateValuesAttribute(PaneStateValues psv)
        {
            var result = "split";
            switch (psv)
            {
                case PaneStateValues.Frozen:
                    result = "frozen";
                    break;
                case PaneStateValues.FrozenSplit:
                    result = "frozenSplit";
                    break;
                case PaneStateValues.Split:
                    result = "split";
                    break;
            }

            return result;
        }
    }
}