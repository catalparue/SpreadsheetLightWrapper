using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.pivottable
{
    internal class SLPivotHierarchy
    {
        internal SLPivotHierarchy()
        {
            SetAllNull();
        }

        internal List<SLMemberProperty> MemberProperties { get; set; }
        internal List<SLMembers> Members { get; set; }

        internal bool Outline { get; set; }
        internal bool MultipleItemSelectionAllowed { get; set; }
        internal bool SubtotalTop { get; set; }
        internal bool ShowInFieldList { get; set; }
        internal bool DragToRow { get; set; }
        internal bool DragToColumn { get; set; }
        internal bool DragToPage { get; set; }
        internal bool DragToData { get; set; }
        internal bool DragOff { get; set; }
        internal bool IncludeNewItemsInFilter { get; set; }
        internal string Caption { get; set; }

        private void SetAllNull()
        {
            MemberProperties = new List<SLMemberProperty>();
            Members = new List<SLMembers>();

            Outline = false;
            MultipleItemSelectionAllowed = false;
            SubtotalTop = false;
            ShowInFieldList = true;
            DragToRow = true;
            DragToColumn = true;
            DragToPage = true;
            DragToData = false;
            DragOff = true;
            IncludeNewItemsInFilter = false;
            Caption = "";
        }

        internal void FromPivotHierarchy(PivotHierarchy ph)
        {
            SetAllNull();

            if (ph.Outline != null) Outline = ph.Outline.Value;
            if (ph.MultipleItemSelectionAllowed != null) Outline = ph.MultipleItemSelectionAllowed.Value;
            if (ph.SubtotalTop != null) SubtotalTop = ph.SubtotalTop.Value;
            if (ph.ShowInFieldList != null) ShowInFieldList = ph.ShowInFieldList.Value;
            if (ph.DragToRow != null) DragToRow = ph.DragToRow.Value;
            if (ph.DragToColumn != null) DragToColumn = ph.DragToColumn.Value;
            if (ph.DragToPage != null) DragToPage = ph.DragToPage.Value;
            if (ph.DragToData != null) DragToData = ph.DragToData.Value;
            if (ph.DragOff != null) DragOff = ph.DragOff.Value;
            if (ph.IncludeNewItemsInFilter != null) IncludeNewItemsInFilter = ph.IncludeNewItemsInFilter.Value;
            if (ph.Caption != null) Caption = ph.Caption.Value;

            SLMemberProperty mp;
            SLMembers mems;
            using (var oxr = OpenXmlReader.Create(ph))
            {
                while (oxr.Read())
                    if (oxr.ElementType == typeof(MemberProperty))
                    {
                        mp = new SLMemberProperty();
                        mp.FromMemberProperty((MemberProperty) oxr.LoadCurrentElement());
                        MemberProperties.Add(mp);
                    }
                    else if (oxr.ElementType == typeof(Members))
                    {
                        mems = new SLMembers();
                        mems.FromMembers((Members) oxr.LoadCurrentElement());
                        Members.Add(mems);
                    }
            }
        }

        internal PivotHierarchy ToPivotHierarchy()
        {
            var ph = new PivotHierarchy();

            if (Outline) ph.Outline = Outline;
            if (MultipleItemSelectionAllowed) ph.MultipleItemSelectionAllowed = MultipleItemSelectionAllowed;
            if (SubtotalTop) ph.SubtotalTop = SubtotalTop;
            if (ShowInFieldList != true) ph.ShowInFieldList = ShowInFieldList;
            if (DragToRow != true) ph.DragToRow = DragToRow;
            if (DragToColumn != true) ph.DragToColumn = DragToColumn;
            if (DragToPage != true) ph.DragToPage = DragToPage;
            if (DragToData) ph.DragToData = DragToData;
            if (DragOff != true) ph.DragOff = DragOff;
            if (IncludeNewItemsInFilter) ph.IncludeNewItemsInFilter = IncludeNewItemsInFilter;
            if ((Caption != null) && (Caption.Length > 0)) ph.Caption = Caption;

            if (MemberProperties.Count > 0)
            {
                ph.MemberProperties = new MemberProperties {Count = (uint) MemberProperties.Count};
                foreach (var mp in MemberProperties)
                    ph.MemberProperties.Append(mp.ToMemberProperty());
            }

            foreach (var mems in Members)
                ph.Append(mems.ToMembers());

            return ph;
        }

        internal SLPivotHierarchy Clone()
        {
            var ph = new SLPivotHierarchy();
            ph.Outline = Outline;
            ph.MultipleItemSelectionAllowed = MultipleItemSelectionAllowed;
            ph.SubtotalTop = SubtotalTop;
            ph.ShowInFieldList = ShowInFieldList;
            ph.DragToRow = DragToRow;
            ph.DragToColumn = DragToColumn;
            ph.DragToPage = DragToPage;
            ph.DragToData = DragToData;
            ph.DragOff = DragOff;
            ph.IncludeNewItemsInFilter = IncludeNewItemsInFilter;
            ph.Caption = Caption;

            ph.MemberProperties = new List<SLMemberProperty>();
            foreach (var mp in MemberProperties)
                ph.MemberProperties.Add(mp.Clone());

            ph.Members = new List<SLMembers>();
            foreach (var mems in Members)
                ph.Members.Add(mems.Clone());

            return ph;
        }
    }
}