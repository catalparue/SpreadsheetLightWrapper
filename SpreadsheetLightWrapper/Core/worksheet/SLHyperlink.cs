using System;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLightWrapper.Core.misc;

namespace SpreadsheetLightWrapper.Core.worksheet
{
    internal class SLHyperlink
    {
        internal string HyperlinkUri;
        internal UriKind HyperlinkUriKind;
        internal bool IsExternal;
        internal bool IsNew;

        internal SLHyperlink()
        {
            SetAllNull();
        }

        internal SLCellPointRange Reference { get; set; }
        internal string Id { get; set; }
        internal string Location { get; set; }
        internal string ToolTip { get; set; }
        internal string Display { get; set; }

        private void SetAllNull()
        {
            IsExternal = false;
            HyperlinkUri = string.Empty;
            HyperlinkUriKind = UriKind.RelativeOrAbsolute;
            IsNew = true;

            Reference = new SLCellPointRange();
            Id = null;
            Location = null;
            ToolTip = null;
            Display = null;
        }

        internal void FromHyperlink(Hyperlink hl)
        {
            SetAllNull();

            IsNew = false;

            if (hl.Reference != null) Reference = SLTool.TranslateReferenceToCellPointRange(hl.Reference.Value);

            if (hl.Id != null)
            {
                // At least I think if there's a relationship ID, it's an external link.
                IsExternal = true;
                Id = hl.Id.Value;
            }

            if (hl.Location != null) Location = hl.Location.Value;
            if (hl.Tooltip != null) ToolTip = hl.Tooltip.Value;
            if (hl.Display != null) Display = hl.Display.Value;
        }

        internal Hyperlink ToHyperlink()
        {
            var hl = new Hyperlink();
            hl.Reference = SLTool.ToCellRange(Reference.StartRowIndex, Reference.StartColumnIndex, Reference.EndRowIndex,
                Reference.EndColumnIndex);
            if ((Id != null) && (Id.Length > 0)) hl.Id = Id;
            if ((Location != null) && (Location.Length > 0)) hl.Location = Location;
            if ((ToolTip != null) && (ToolTip.Length > 0)) hl.Tooltip = ToolTip;
            if ((Display != null) && (Display.Length > 0)) hl.Display = Display;

            return hl;
        }

        internal SLHyperlink Clone()
        {
            var hl = new SLHyperlink();
            hl.IsExternal = IsExternal;
            hl.HyperlinkUri = HyperlinkUri;
            hl.HyperlinkUriKind = HyperlinkUriKind;
            hl.IsNew = IsNew;
            hl.Reference = new SLCellPointRange(Reference.StartRowIndex, Reference.StartColumnIndex,
                Reference.EndRowIndex, Reference.EndColumnIndex);
            hl.Id = Id;
            hl.Location = Location;
            hl.ToolTip = ToolTip;
            hl.Display = Display;

            return hl;
        }
    }
}