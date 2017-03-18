using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.pivottable
{
    internal class SLMembers
    {
        internal SLMembers()
        {
            SetAllNull();
        }

        internal List<string> Members { get; set; }
        internal uint? Level { get; set; }

        private void SetAllNull()
        {
            Members = new List<string>();
            Level = null;
        }

        internal void FromMembers(Members m)
        {
            SetAllNull();

            if (m.Level != null) Level = m.Level.Value;

            Member mem;
            using (var oxr = OpenXmlReader.Create(m))
            {
                while (oxr.Read())
                    if (oxr.ElementType == typeof(Member))
                    {
                        mem = (Member) oxr.LoadCurrentElement();
                        Members.Add(mem.Name.Value);
                    }
            }
        }

        internal Members ToMembers()
        {
            var m = new Members();
            m.Count = (uint) Members.Count;
            if (Level != null) m.Level = Level.Value;

            foreach (var s in Members)
                m.Append(new Member {Name = s});

            return m;
        }

        internal SLMembers Clone()
        {
            var m = new SLMembers();
            m.Level = Level;

            m.Members = new List<string>();
            foreach (var s in Members)
                m.Members.Add(s);

            return m;
        }
    }
}