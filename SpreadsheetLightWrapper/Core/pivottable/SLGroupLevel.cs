using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.pivottable
{
    internal class SLGroupLevel
    {
        internal SLGroupLevel()
        {
            SetAllNull();
        }

        internal List<SLGroup> Groups { get; set; }

        internal string UniqueName { get; set; }
        internal string Caption { get; set; }
        internal bool User { get; set; }
        internal bool CustomRollUp { get; set; }

        private void SetAllNull()
        {
            Groups = new List<SLGroup>();

            UniqueName = "";
            Caption = "";
            User = false;
            CustomRollUp = false;
        }

        internal void FromGroupLevel(GroupLevel gl)
        {
            SetAllNull();

            if (gl.UniqueName != null) UniqueName = gl.UniqueName.Value;
            if (gl.Caption != null) Caption = gl.Caption.Value;
            if (gl.User != null) User = gl.User.Value;
            if (gl.CustomRollUp != null) CustomRollUp = gl.CustomRollUp.Value;

            if (gl.Groups != null)
            {
                SLGroup g;
                using (var oxr = OpenXmlReader.Create(gl.Groups))
                {
                    while (oxr.Read())
                        if (oxr.ElementType == typeof(Group))
                        {
                            g = new SLGroup();
                            g.FromGroup((Group) oxr.LoadCurrentElement());
                            Groups.Add(g);
                        }
                }
            }
        }

        internal GroupLevel ToGroupLevel()
        {
            var gl = new GroupLevel();
            gl.UniqueName = UniqueName;
            gl.Caption = Caption;
            if (User) gl.User = User;
            if (CustomRollUp) gl.CustomRollUp = CustomRollUp;

            if (Groups.Count > 0)
            {
                gl.Groups = new Groups {Count = (uint) Groups.Count};
                foreach (var g in Groups)
                    gl.Groups.Append(g.ToGroup());
            }

            return gl;
        }

        internal SLGroupLevel Clone()
        {
            var gl = new SLGroupLevel();
            gl.UniqueName = UniqueName;
            gl.Caption = Caption;
            gl.User = User;
            gl.CustomRollUp = CustomRollUp;

            gl.Groups = new List<SLGroup>();
            foreach (var g in Groups)
                gl.Groups.Add(g.Clone());

            return gl;
        }
    }
}