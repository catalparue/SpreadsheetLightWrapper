using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.pivottable
{
    internal class SLRangeProperties
    {
        internal SLRangeProperties()
        {
            SetAllNull();
        }

        internal bool AutoStart { get; set; }
        internal bool AutoEnd { get; set; }
        internal GroupByValues GroupBy { get; set; }
        internal double? StartNumber { get; set; }
        internal double? EndNum { get; set; }
        internal DateTime? StartDate { get; set; }
        internal DateTime? EndDate { get; set; }
        internal double GroupInterval { get; set; }

        private void SetAllNull()
        {
            AutoStart = true;
            AutoEnd = true;
            GroupBy = GroupByValues.Range;
            StartNumber = null;
            EndNum = null;
            StartDate = null;
            EndDate = null;
            GroupInterval = 1;
        }

        internal void FromRangeProperties(RangeProperties rp)
        {
            SetAllNull();

            if (rp.AutoStart != null) AutoStart = rp.AutoStart.Value;
            if (rp.AutoEnd != null) AutoEnd = rp.AutoEnd.Value;
            if (rp.GroupBy != null) GroupBy = rp.GroupBy.Value;
            if (rp.StartNumber != null) StartNumber = rp.StartNumber.Value;
            if (rp.EndNum != null) EndNum = rp.EndNum.Value;
            if (rp.StartDate != null) StartDate = rp.StartDate.Value;
            if (rp.EndDate != null) EndDate = rp.EndDate.Value;
            if (rp.GroupInterval != null) GroupInterval = rp.GroupInterval.Value;
        }

        internal RangeProperties ToRangeProperties()
        {
            var rp = new RangeProperties();
            if (AutoStart != true) rp.AutoStart = AutoStart;
            if (AutoEnd != true) rp.AutoEnd = AutoEnd;
            if (GroupBy != GroupByValues.Range) rp.GroupBy = GroupBy;
            if (StartNumber != null) rp.StartNumber = StartNumber.Value;
            if (EndNum != null) rp.EndNum = EndNum.Value;
            if (StartDate != null) rp.StartDate = StartDate.Value;
            if (EndDate != null) rp.EndDate = EndDate.Value;
            if (GroupInterval != 1) rp.GroupInterval = GroupInterval;

            return rp;
        }

        internal SLRangeProperties Clone()
        {
            var rp = new SLRangeProperties();
            rp.AutoStart = AutoStart;
            rp.AutoEnd = AutoEnd;
            rp.GroupBy = GroupBy;
            rp.StartNumber = StartNumber;
            rp.EndNum = EndNum;
            rp.StartDate = StartDate;
            rp.EndDate = EndDate;
            rp.GroupInterval = GroupInterval;

            return rp;
        }
    }
}