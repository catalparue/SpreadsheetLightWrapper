using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.pivottable
{
    internal class SLConsolidation
    {
        internal SLConsolidation()
        {
            SetAllNull();
        }

        internal List<List<string>> Pages { get; set; }
        internal List<SLRangeSet> RangeSets { get; set; }

        internal bool AutoPage { get; set; }

        private void SetAllNull()
        {
            Pages = new List<List<string>>();
            RangeSets = new List<SLRangeSet>();
            AutoPage = true;
        }

        internal void FromConsolidation(Consolidation c)
        {
            SetAllNull();

            if (c.AutoPage != null) AutoPage = c.AutoPage.Value;

            Page pg;
            PageItem pgi;
            List<string> listPage;
            SLRangeSet rs;
            using (var oxr = OpenXmlReader.Create(c))
            {
                while (oxr.Read())
                    if (oxr.ElementType == typeof(Page))
                    {
                        listPage = new List<string>();
                        pg = (Page) oxr.LoadCurrentElement();
                        using (var oxrPage = OpenXmlReader.Create(pg))
                        {
                            while (oxrPage.Read())
                                if (oxrPage.ElementType == typeof(PageItem))
                                {
                                    pgi = (PageItem) oxrPage.LoadCurrentElement();
                                    listPage.Add(pgi.Name.Value);
                                }
                        }
                        Pages.Add(listPage);
                    }
                    else if (oxr.ElementType == typeof(RangeSet))
                    {
                        rs = new SLRangeSet();
                        rs.FromRangeSet((RangeSet) oxr.LoadCurrentElement());
                        RangeSets.Add(rs);
                    }
            }
        }

        internal Consolidation ToConsolidation()
        {
            var c = new Consolidation();
            if (AutoPage != true) c.AutoPage = AutoPage;

            if (Pages.Count > 0)
            {
                Page pg;
                c.Pages = new Pages {Count = (uint) Pages.Count};
                foreach (var ls in Pages)
                {
                    pg = new Page {Count = (uint) ls.Count};
                    foreach (var s in ls)
                        pg.Append(new PageItem {Name = s});
                    c.Pages.Append(pg);
                }
            }

            c.RangeSets = new RangeSets {Count = (uint) RangeSets.Count};
            foreach (var rs in RangeSets)
                c.RangeSets.Append(rs.ToRangeSet());

            return c;
        }

        internal SLConsolidation Clone()
        {
            var c = new SLConsolidation();
            c.AutoPage = AutoPage;

            List<string> list;
            foreach (var ls in Pages)
            {
                list = new List<string>();
                foreach (var s in ls)
                    list.Add(s);
                c.Pages.Add(list);
            }

            c.RangeSets = new List<SLRangeSet>();
            foreach (var rs in RangeSets)
                c.RangeSets.Add(rs.Clone());

            return c;
        }
    }
}