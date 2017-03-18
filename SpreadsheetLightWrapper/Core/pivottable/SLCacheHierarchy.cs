using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.pivottable
{
    internal class SLCacheHierarchy
    {
        internal SLCacheHierarchy()
        {
            SetAllNull();
        }

        internal List<int> FieldsUsage { get; set; }
        internal List<SLGroupLevel> GroupLevels { get; set; }

        internal string UniqueName { get; set; }
        internal string Caption { get; set; }
        internal bool Measure { get; set; }
        internal bool Set { get; set; }
        internal uint? ParentSet { get; set; }
        internal int IconSet { get; set; }
        internal bool Attribute { get; set; }
        internal bool Time { get; set; }
        internal bool KeyAttribute { get; set; }
        internal string DefaultMemberUniqueName { get; set; }
        internal string AllUniqueName { get; set; }
        internal string AllCaption { get; set; }
        internal string DimensionUniqueName { get; set; }
        internal string DisplayFolder { get; set; }
        internal string MeasureGroup { get; set; }
        internal bool Measures { get; set; }
        internal uint Count { get; set; }
        internal bool OneField { get; set; }
        internal ushort? MemberValueDatatype { get; set; }
        internal bool? Unbalanced { get; set; }
        internal bool? UnbalancedGroup { get; set; }
        internal bool Hidden { get; set; }

        private void SetAllNull()
        {
            FieldsUsage = new List<int>();
            GroupLevels = new List<SLGroupLevel>();

            UniqueName = "";
            Caption = "";
            Measure = false;
            Set = false;
            ParentSet = null;
            IconSet = 0;
            Attribute = false;
            Time = false;
            KeyAttribute = false;
            DefaultMemberUniqueName = "";
            AllUniqueName = "";
            AllCaption = "";
            DimensionUniqueName = "";
            DisplayFolder = "";
            MeasureGroup = "";
            Measures = false;
            Count = 0;
            OneField = false;
            MemberValueDatatype = null;
            Unbalanced = null;
            UnbalancedGroup = null;
            Hidden = false;
        }

        internal void FromCacheHierarchy(CacheHierarchy ch)
        {
            SetAllNull();

            if (ch.UniqueName != null) UniqueName = ch.UniqueName.Value;
            if (ch.Caption != null) Caption = ch.Caption.Value;
            if (ch.Measure != null) Measure = ch.Measure.Value;
            if (ch.Set != null) Set = ch.Set.Value;
            if (ch.ParentSet != null) ParentSet = ch.ParentSet.Value;
            if (ch.IconSet != null) IconSet = ch.IconSet.Value;
            if (ch.Attribute != null) Attribute = ch.Attribute.Value;
            if (ch.Time != null) Time = ch.Time.Value;
            if (ch.KeyAttribute != null) KeyAttribute = ch.KeyAttribute.Value;
            if (ch.DefaultMemberUniqueName != null) DefaultMemberUniqueName = ch.DefaultMemberUniqueName.Value;
            if (ch.AllUniqueName != null) AllUniqueName = ch.AllUniqueName.Value;
            if (ch.AllCaption != null) AllCaption = ch.AllCaption.Value;
            if (ch.DimensionUniqueName != null) DimensionUniqueName = ch.DimensionUniqueName.Value;
            if (ch.DisplayFolder != null) DisplayFolder = ch.DisplayFolder.Value;
            if (ch.MeasureGroup != null) MeasureGroup = ch.MeasureGroup.Value;
            if (ch.Measures != null) Measures = ch.Measures.Value;
            if (ch.Count != null) Count = ch.Count.Value;
            if (ch.OneField != null) OneField = ch.OneField.Value;
            if (ch.MemberValueDatatype != null) MemberValueDatatype = ch.MemberValueDatatype.Value;
            if (ch.Unbalanced != null) Unbalanced = ch.Unbalanced.Value;
            if (ch.UnbalancedGroup != null) UnbalancedGroup = ch.UnbalancedGroup.Value;
            if (ch.Hidden != null) Hidden = ch.Hidden.Value;

            FieldUsage fu;
            SLGroupLevel gl;
            using (var oxr = OpenXmlReader.Create(ch))
            {
                while (oxr.Read())
                    if (oxr.ElementType == typeof(FieldUsage))
                    {
                        fu = (FieldUsage) oxr.LoadCurrentElement();
                        FieldsUsage.Add(fu.Index.Value);
                    }
                    else if (oxr.ElementType == typeof(GroupLevel))
                    {
                        gl = new SLGroupLevel();
                        gl.FromGroupLevel((GroupLevel) oxr.LoadCurrentElement());
                        GroupLevels.Add(gl);
                    }
            }
        }

        internal CacheHierarchy ToCacheHierarchy()
        {
            var ch = new CacheHierarchy();
            ch.UniqueName = UniqueName;
            if ((Caption != null) && (Caption.Length > 0)) ch.Caption = Caption;
            if (Measure) ch.Measure = Measure;
            if (Set) ch.Set = Set;
            if (ParentSet != null) ch.ParentSet = ParentSet.Value;
            if (IconSet != 0) ch.IconSet = IconSet;
            if (Attribute) ch.Attribute = Attribute;
            if (Time) ch.Time = Time;
            if (KeyAttribute) ch.KeyAttribute = KeyAttribute;
            if ((DefaultMemberUniqueName != null) && (DefaultMemberUniqueName.Length > 0))
                ch.DefaultMemberUniqueName = DefaultMemberUniqueName;
            if ((AllUniqueName != null) && (AllUniqueName.Length > 0)) ch.AllUniqueName = AllUniqueName;
            if ((AllCaption != null) && (AllCaption.Length > 0)) ch.AllCaption = AllCaption;
            if ((DimensionUniqueName != null) && (DimensionUniqueName.Length > 0))
                ch.DimensionUniqueName = DimensionUniqueName;
            if ((DisplayFolder != null) && (DisplayFolder.Length > 0)) ch.DisplayFolder = DisplayFolder;
            if ((MeasureGroup != null) && (MeasureGroup.Length > 0)) ch.MeasureGroup = MeasureGroup;
            if (Measures) ch.Measures = Measures;
            ch.Count = Count;
            if (OneField) ch.OneField = OneField;
            if (MemberValueDatatype != null) ch.MemberValueDatatype = MemberValueDatatype.Value;
            if (Unbalanced != null) ch.Unbalanced = Unbalanced.Value;
            if (UnbalancedGroup != null) ch.UnbalancedGroup = UnbalancedGroup.Value;
            if (Hidden) ch.Hidden = Hidden;

            if (FieldsUsage.Count > 0)
            {
                ch.FieldsUsage = new FieldsUsage {Count = (uint) FieldsUsage.Count};
                foreach (var i in FieldsUsage)
                    ch.FieldsUsage.Append(new FieldUsage {Index = i});
            }

            if (GroupLevels.Count > 0)
            {
                ch.GroupLevels = new GroupLevels {Count = (uint) GroupLevels.Count};
                foreach (var gl in GroupLevels)
                    ch.GroupLevels.Append(gl.ToGroupLevel());
            }

            return ch;
        }

        internal SLCacheHierarchy Clone()
        {
            var ch = new SLCacheHierarchy();
            ch.UniqueName = UniqueName;
            ch.Caption = Caption;
            ch.Measure = Measure;
            ch.Set = Set;
            ch.ParentSet = ParentSet;
            ch.IconSet = IconSet;
            ch.Attribute = Attribute;
            ch.Time = Time;
            ch.KeyAttribute = KeyAttribute;
            ch.DefaultMemberUniqueName = DefaultMemberUniqueName;
            ch.AllUniqueName = AllUniqueName;
            ch.AllCaption = AllCaption;
            ch.DimensionUniqueName = DimensionUniqueName;
            ch.DisplayFolder = DisplayFolder;
            ch.MeasureGroup = MeasureGroup;
            ch.Measures = Measures;
            ch.Count = Count;
            ch.OneField = OneField;
            ch.MemberValueDatatype = MemberValueDatatype;
            ch.Unbalanced = Unbalanced;
            ch.UnbalancedGroup = UnbalancedGroup;
            ch.Hidden = Hidden;

            ch.FieldsUsage = new List<int>();
            foreach (var i in FieldsUsage)
                ch.FieldsUsage.Add(i);

            ch.GroupLevels = new List<SLGroupLevel>();
            foreach (var gl in GroupLevels)
                ch.GroupLevels.Add(gl.Clone());

            return ch;
        }
    }
}