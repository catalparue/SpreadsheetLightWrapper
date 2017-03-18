using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.table
{
    internal class SLDateGroupItem
    {
        internal SLDateGroupItem()
        {
            SetAllNull();
        }

        internal ushort Year { get; set; }
        internal ushort? Month { get; set; }
        internal ushort? Day { get; set; }
        internal ushort? Hour { get; set; }
        internal ushort? Minute { get; set; }
        internal ushort? Second { get; set; }
        internal DateTimeGroupingValues DateTimeGrouping { get; set; }

        private void SetAllNull()
        {
            Year = (ushort) DateTime.Now.Year;
            Month = null;
            Day = null;
            Hour = null;
            Minute = null;
            Second = null;
            DateTimeGrouping = DateTimeGroupingValues.Year;
        }

        internal void FromDateGroupItem(DateGroupItem dgi)
        {
            SetAllNull();

            Year = dgi.Year.Value;
            if (dgi.Month != null) Month = dgi.Month.Value;
            if (dgi.Day != null) Day = dgi.Day.Value;
            if (dgi.Hour != null) Hour = dgi.Hour.Value;
            if (dgi.Minute != null) Minute = dgi.Minute.Value;
            if (dgi.Second != null) Second = dgi.Second.Value;
            DateTimeGrouping = dgi.DateTimeGrouping.Value;
        }

        internal DateGroupItem ToDateGroupItem()
        {
            var dgi = new DateGroupItem();
            dgi.Year = Year;
            if (Month != null) dgi.Month = Month.Value;
            if (Day != null) dgi.Day = Day.Value;
            if (Hour != null) dgi.Hour = Hour.Value;
            if (Minute != null) dgi.Minute = Minute.Value;
            if (Second != null) dgi.Second = Second.Value;
            dgi.DateTimeGrouping = DateTimeGrouping;

            return dgi;
        }

        internal SLDateGroupItem Clone()
        {
            var dgi = new SLDateGroupItem();
            dgi.Year = Year;
            dgi.Month = Month;
            dgi.Day = Day;
            dgi.Hour = Hour;
            dgi.Minute = Minute;
            dgi.Second = Second;
            dgi.DateTimeGrouping = DateTimeGrouping;

            return dgi;
        }
    }
}