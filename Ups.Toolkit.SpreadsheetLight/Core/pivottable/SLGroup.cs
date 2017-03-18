using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.pivottable
{
    internal class SLGroup
    {
        internal SLGroup()
        {
            SetAllNull();
        }

        internal List<SLGroupMember> GroupMembers { get; set; }

        internal string Name { get; set; }
        internal string UniqueName { get; set; }
        internal string Caption { get; set; }
        internal string UniqueParent { get; set; }
        internal int? Id { get; set; }

        private void SetAllNull()
        {
            GroupMembers = new List<SLGroupMember>();

            Name = "";
            UniqueName = "";
            Caption = "";
            UniqueParent = "";
            Id = null;
        }

        internal void FromGroup(Group g)
        {
            SetAllNull();

            if (g.Name != null) Name = g.Name.Value;
            if (g.UniqueName != null) UniqueName = g.UniqueName.Value;
            if (g.Caption != null) Caption = g.Caption.Value;
            if (g.UniqueParent != null) UniqueParent = g.UniqueParent.Value;
            if (g.Id != null) Id = g.Id.Value;
        }

        internal Group ToGroup()
        {
            var g = new Group();
            g.Name = Name;
            g.UniqueName = UniqueName;
            g.Caption = Caption;
            if ((UniqueParent != null) && (UniqueParent.Length > 0)) g.UniqueParent = UniqueParent;
            if (Id != null) g.Id = Id.Value;

            if (GroupMembers.Count > 0)
            {
                g.GroupMembers = new GroupMembers {Count = (uint) GroupMembers.Count};
                foreach (var gm in GroupMembers)
                    g.GroupMembers.Append(gm.ToGroupMember());
            }

            return g;
        }

        internal SLGroup Clone()
        {
            var g = new SLGroup();
            g.Name = Name;
            g.UniqueName = UniqueName;
            g.Caption = Caption;
            g.UniqueParent = UniqueParent;
            g.Id = Id;

            g.GroupMembers = new List<SLGroupMember>();
            foreach (var gm in GroupMembers)
                g.GroupMembers.Add(gm.Clone());

            return g;
        }
    }
}