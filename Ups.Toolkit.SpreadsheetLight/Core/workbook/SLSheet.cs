using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.workbook
{
    internal class SLSheet
    {
        internal SLSheet(string Name, uint SheetId, string Id, SLSheetType SheetType)
        {
            this.Name = Name;
            this.SheetId = SheetId;
            State = SheetStateValues.Visible;
            this.Id = Id;
            this.SheetType = SheetType;
        }

        internal string Name { get; set; }
        internal uint SheetId { get; set; }
        internal SheetStateValues State { get; set; }
        internal string Id { get; set; }
        internal SLSheetType SheetType { get; set; }

        internal Sheet ToSheet()
        {
            var s = new Sheet();
            s.Name = Name;
            s.SheetId = SheetId;
            if (State != SheetStateValues.Visible) s.State = State;
            s.Id = Id;

            return s;
        }

        internal SLSheet Clone()
        {
            var s = new SLSheet(Name, SheetId, Id, SheetType);
            s.State = State;
            return s;
        }
    }
}