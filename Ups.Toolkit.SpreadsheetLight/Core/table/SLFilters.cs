using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.table
{
    internal class SLFilters
    {
        internal bool HasCalendarType;
        private CalendarValues vCalendarType;

        internal SLFilters()
        {
            SetAllNull();
        }

        internal List<SLFilter> Filters { get; set; }
        internal List<SLDateGroupItem> DateGroupItems { get; set; }
        internal bool? Blank { get; set; }

        internal CalendarValues CalendarType
        {
            get { return vCalendarType; }
            set
            {
                vCalendarType = value;
                HasCalendarType = vCalendarType != CalendarValues.None ? true : false;
            }
        }

        private void SetAllNull()
        {
            Filters = new List<SLFilter>();
            DateGroupItems = new List<SLDateGroupItem>();
            Blank = null;
            vCalendarType = CalendarValues.None;
            HasCalendarType = false;
        }

        internal void FromFilters(Filters fs)
        {
            SetAllNull();

            if ((fs.Blank != null) && fs.Blank.Value) Blank = fs.Blank.Value;
            if (fs.CalendarType != null) CalendarType = fs.CalendarType.Value;

            if (fs.HasChildren)
            {
                SLFilter f;
                SLDateGroupItem dgi;
                using (var oxr = OpenXmlReader.Create(fs))
                {
                    while (oxr.Read())
                        if (oxr.ElementType == typeof(Filter))
                        {
                            f = new SLFilter();
                            f.FromFilter((Filter) oxr.LoadCurrentElement());
                            Filters.Add(f);
                        }
                        else if (oxr.ElementType == typeof(DateGroupItem))
                        {
                            dgi = new SLDateGroupItem();
                            dgi.FromDateGroupItem((DateGroupItem) oxr.LoadCurrentElement());
                            DateGroupItems.Add(dgi);
                        }
                }
            }
        }

        internal Filters ToFilters()
        {
            var fs = new Filters();
            if ((Blank != null) && Blank.Value) fs.Blank = Blank.Value;
            if (HasCalendarType) fs.CalendarType = CalendarType;

            foreach (var f in Filters)
                fs.Append(f.ToFilter());

            foreach (var dgi in DateGroupItems)
                fs.Append(dgi.ToDateGroupItem());

            return fs;
        }

        internal SLFilters Clone()
        {
            var fs = new SLFilters();

            int i;
            fs.Filters = new List<SLFilter>();
            for (i = 0; i < Filters.Count; ++i)
                fs.Filters.Add(Filters[i].Clone());

            fs.DateGroupItems = new List<SLDateGroupItem>();
            for (i = 0; i < DateGroupItems.Count; ++i)
                fs.DateGroupItems.Add(DateGroupItems[i].Clone());

            fs.Blank = Blank;
            fs.HasCalendarType = HasCalendarType;
            fs.vCalendarType = vCalendarType;

            return fs;
        }
    }
}