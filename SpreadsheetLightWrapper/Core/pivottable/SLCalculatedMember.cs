using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.pivottable
{
    internal class SLCalculatedMember
    {
        internal SLCalculatedMember()
        {
            SetAllNull();
        }

        internal string Name { get; set; }
        internal string Mdx { get; set; }
        internal string MemberName { get; set; }
        internal string Hierarchy { get; set; }
        internal string ParentName { get; set; }
        internal int SolveOrder { get; set; }
        internal bool Set { get; set; }

        private void SetAllNull()
        {
            Name = "";
            Mdx = "";
            MemberName = "";
            Hierarchy = "";
            ParentName = "";
            SolveOrder = 0;
            Set = false;
        }

        internal void FromCalculatedMember(CalculatedMember cm)
        {
            SetAllNull();

            if (cm.Name != null) Name = cm.Name.Value;
            if (cm.Mdx != null) Mdx = cm.Mdx.Value;
            if (cm.MemberName != null) MemberName = cm.MemberName.Value;
            if (cm.Hierarchy != null) Hierarchy = cm.Hierarchy.Value;
            if (cm.ParentName != null) ParentName = cm.ParentName.Value;
            if (cm.SolveOrder != null) SolveOrder = cm.SolveOrder.Value;
            if (cm.Set != null) Set = cm.Set.Value;
        }

        internal CalculatedMember ToCalculatedMember()
        {
            var cm = new CalculatedMember();
            cm.Name = Name;
            cm.Mdx = Mdx;
            if ((MemberName != null) && (MemberName.Length > 0)) cm.MemberName = MemberName;
            if ((Hierarchy != null) && (Hierarchy.Length > 0)) cm.Hierarchy = Hierarchy;
            if ((ParentName != null) && (ParentName.Length > 0)) cm.ParentName = ParentName;
            if (SolveOrder != 0) cm.SolveOrder = SolveOrder;
            if (Set) cm.Set = Set;

            return cm;
        }

        internal SLCalculatedMember Clone()
        {
            var cm = new SLCalculatedMember();
            cm.Name = Name;
            cm.Mdx = Mdx;
            cm.MemberName = MemberName;
            cm.Hierarchy = Hierarchy;
            cm.ParentName = ParentName;
            cm.SolveOrder = SolveOrder;
            cm.Set = Set;

            return cm;
        }
    }
}