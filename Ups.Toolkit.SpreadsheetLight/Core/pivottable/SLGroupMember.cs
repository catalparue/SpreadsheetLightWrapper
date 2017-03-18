using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.pivottable
{
    internal class SLGroupMember
    {
        internal SLGroupMember()
        {
            SetAllNull();
        }

        internal string UniqueName { get; set; }
        internal bool Group { get; set; }

        private void SetAllNull()
        {
            UniqueName = "";
            Group = false;
        }

        internal void FromGroupMember(GroupMember gm)
        {
            SetAllNull();

            if (gm.UniqueName != null) UniqueName = gm.UniqueName.Value;
            if (gm.Group != null) Group = gm.Group.Value;
        }

        internal GroupMember ToGroupMember()
        {
            var gm = new GroupMember();
            gm.UniqueName = UniqueName;
            if (Group) gm.Group = Group;

            return gm;
        }

        internal SLGroupMember Clone()
        {
            var gm = new SLGroupMember();
            gm.UniqueName = UniqueName;
            gm.Group = Group;

            return gm;
        }
    }
}