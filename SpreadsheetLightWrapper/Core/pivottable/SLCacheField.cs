using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.pivottable
{
    internal class SLCacheField
    {
        internal bool HasFieldGroup;
        internal bool HasSharedItems;

        internal SLCacheField()
        {
            SetAllNull();
        }

        internal SLSharedItems SharedItems { get; set; }
        internal SLFieldGroup FieldGroup { get; set; }

        internal List<int> MemberPropertiesMaps { get; set; }

        internal string Name { get; set; }
        internal string Caption { get; set; }
        internal string PropertyName { get; set; }
        internal bool ServerField { get; set; }
        internal bool UniqueList { get; set; }
        internal uint? NumberFormatId { get; set; }
        internal string Formula { get; set; }
        internal int SqlType { get; set; }
        internal int Hierarchy { get; set; }
        internal uint Level { get; set; }
        internal bool DatabaseField { get; set; }
        internal uint? MappingCount { get; set; }
        internal bool MemberPropertyField { get; set; }

        private void SetAllNull()
        {
            HasSharedItems = false;
            SharedItems = new SLSharedItems();

            HasFieldGroup = false;
            FieldGroup = new SLFieldGroup();

            MemberPropertiesMaps = new List<int>();

            Name = "";
            Caption = "";
            PropertyName = "";
            ServerField = false;
            UniqueList = true;
            NumberFormatId = null;
            Formula = "";
            SqlType = 0;
            Hierarchy = 0;
            Level = 0;
            DatabaseField = true;
            MappingCount = null;
            MemberPropertyField = false;
        }

        internal void FromCacheField(CacheField cf)
        {
            SetAllNull();

            if (cf.Name != null) Name = cf.Name.Value;
            if (cf.Caption != null) Caption = cf.Caption.Value;
            if (cf.PropertyName != null) PropertyName = cf.PropertyName.Value;
            if (cf.ServerField != null) ServerField = cf.ServerField.Value;
            if (cf.UniqueList != null) UniqueList = cf.UniqueList.Value;
            if (cf.NumberFormatId != null) NumberFormatId = cf.NumberFormatId.Value;
            if (cf.Formula != null) Formula = cf.Formula.Value;
            if (cf.SqlType != null) SqlType = cf.SqlType.Value;
            if (cf.Hierarchy != null) Hierarchy = cf.Hierarchy.Value;
            if (cf.Level != null) Level = cf.Level.Value;
            if (cf.DatabaseField != null) DatabaseField = cf.DatabaseField.Value;
            if (cf.MappingCount != null) MappingCount = cf.MappingCount.Value;
            if (cf.MemberPropertyField != null) MemberPropertyField = cf.MemberPropertyField.Value;

            if (cf.SharedItems != null)
            {
                SharedItems.FromSharedItems(cf.SharedItems);
                HasSharedItems = true;
            }

            if (cf.FieldGroup != null)
            {
                FieldGroup.FromFieldGroup(cf.FieldGroup);
                HasFieldGroup = true;
            }

            MemberPropertiesMap mpm;
            using (var oxr = OpenXmlReader.Create(cf))
            {
                while (oxr.Read())
                    if (oxr.ElementType == typeof(MemberPropertiesMap))
                    {
                        mpm = (MemberPropertiesMap) oxr.LoadCurrentElement();
                        if (mpm.Val != null) MemberPropertiesMaps.Add(mpm.Val.Value);
                        else MemberPropertiesMaps.Add(0);
                    }
            }
        }

        internal CacheField ToCacheField()
        {
            var cf = new CacheField();
            cf.Name = Name;
            if ((Caption != null) && (Caption.Length > 0)) cf.Caption = Caption;
            if ((PropertyName != null) & (PropertyName.Length > 0)) cf.PropertyName = PropertyName;
            if (ServerField) cf.ServerField = ServerField;
            if (UniqueList != true) cf.UniqueList = UniqueList;
            if (NumberFormatId != null) cf.NumberFormatId = NumberFormatId.Value;
            if ((Formula != null) && (Formula.Length > 0)) cf.Formula = Formula;
            if (SqlType != 0) cf.SqlType = SqlType;
            if (Hierarchy != 0) cf.Hierarchy = Hierarchy;
            if (Level != 0) cf.Level = Level;
            if (DatabaseField != true) cf.DatabaseField = DatabaseField;
            if (MappingCount != null) cf.MappingCount = MappingCount.Value;
            if (MemberPropertyField) cf.MemberPropertyField = MemberPropertyField;

            if (HasSharedItems)
                cf.SharedItems = SharedItems.ToSharedItems();

            if (HasFieldGroup)
                cf.FieldGroup = FieldGroup.ToFieldGroup();

            foreach (var i in MemberPropertiesMaps)
                if (i != 0) cf.Append(new MemberPropertiesMap {Val = i});
                else cf.Append(new MemberPropertiesMap());

            return cf;
        }

        internal SLCacheField Clone()
        {
            var cf = new SLCacheField();
            cf.Name = Name;
            cf.Caption = Caption;
            cf.PropertyName = PropertyName;
            cf.ServerField = ServerField;
            cf.UniqueList = UniqueList;
            cf.NumberFormatId = NumberFormatId;
            cf.Formula = Formula;
            cf.SqlType = SqlType;
            cf.Hierarchy = Hierarchy;
            cf.Level = Level;
            cf.DatabaseField = DatabaseField;
            cf.MappingCount = MappingCount;
            cf.MemberPropertyField = MemberPropertyField;

            cf.HasSharedItems = HasSharedItems;
            cf.SharedItems = SharedItems.Clone();

            cf.HasFieldGroup = HasFieldGroup;
            cf.FieldGroup = FieldGroup.Clone();

            cf.MemberPropertiesMaps = new List<int>();
            foreach (var i in MemberPropertiesMaps)
                cf.MemberPropertiesMaps.Add(i);

            return cf;
        }
    }
}