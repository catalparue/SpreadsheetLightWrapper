using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;
using Ups.Toolkit.SpreadsheetLight.Core.misc;

namespace Ups.Toolkit.SpreadsheetLight.Core.worksheet
{
    internal class SLSelection
    {
        internal SLSelection()
        {
            SetAllNull();
        }

        internal PaneValues Pane { get; set; }
        internal string ActiveCell { get; set; }
        internal uint ActiveCellId { get; set; }
        internal List<SLCellPointRange> SequenceOfReferences { get; set; }

        private void SetAllNull()
        {
            Pane = PaneValues.TopLeft;
            ActiveCell = string.Empty;
            ActiveCellId = 0;
            SequenceOfReferences = new List<SLCellPointRange>();
        }

        internal void FromSelection(Selection s)
        {
            SetAllNull();

            if (s.Pane != null) Pane = s.Pane.Value;
            if (s.ActiveCell != null) ActiveCell = s.ActiveCell.Value;
            if (s.ActiveCellId != null) ActiveCellId = s.ActiveCellId.Value;
            if (s.SequenceOfReferences != null)
                SequenceOfReferences = SLTool.TranslateSeqRefToCellPointRange(s.SequenceOfReferences);
        }

        internal Selection ToSelection()
        {
            var s = new Selection();
            if (Pane != PaneValues.TopLeft) s.Pane = Pane;

            if ((ActiveCell.Length > 0) && !ActiveCell.Equals("A1", StringComparison.OrdinalIgnoreCase))
                s.ActiveCell = ActiveCell;

            if (ActiveCellId != 0) s.ActiveCellId = ActiveCellId;

            if (SequenceOfReferences.Count > 0)
                if (SequenceOfReferences.Count == 1)
                {
                    // not equal to A1
                    if ((SequenceOfReferences[0].StartRowIndex != 1)
                        || (SequenceOfReferences[0].StartColumnIndex != 1)
                        || (SequenceOfReferences[0].EndRowIndex != 1)
                        || (SequenceOfReferences[0].EndColumnIndex != 1))
                        s.SequenceOfReferences = SLTool.TranslateCellPointRangeToSeqRef(SequenceOfReferences);
                }
                else
                {
                    s.SequenceOfReferences = SLTool.TranslateCellPointRangeToSeqRef(SequenceOfReferences);
                }

            return s;
        }

        internal SLSelection Clone()
        {
            var s = new SLSelection();
            s.Pane = Pane;
            s.ActiveCell = ActiveCell;
            s.ActiveCellId = ActiveCellId;

            s.SequenceOfReferences = new List<SLCellPointRange>();
            foreach (var pt in SequenceOfReferences)
                s.SequenceOfReferences.Add(new SLCellPointRange(pt.StartRowIndex, pt.StartColumnIndex, pt.EndRowIndex,
                    pt.EndColumnIndex));

            return s;
        }
    }
}