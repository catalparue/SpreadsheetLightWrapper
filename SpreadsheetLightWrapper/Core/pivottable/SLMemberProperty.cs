using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetLightWrapper.Core.pivottable
{
    internal class SLMemberProperty
    {
        internal SLMemberProperty()
        {
            SetAllNull();
        }

        internal string Name { get; set; }
        internal bool ShowCell { get; set; }
        internal bool ShowTip { get; set; }
        internal bool ShowAsCaption { get; set; }
        internal uint? NameLength { get; set; }
        internal uint? PropertyNamePosition { get; set; }
        internal uint? PropertyNameLength { get; set; }
        internal uint? Level { get; set; }
        internal uint Field { get; set; }

        private void SetAllNull()
        {
            Name = "";
            ShowCell = false;
            ShowTip = false;
            ShowAsCaption = false;
            NameLength = null;
            PropertyNamePosition = null;
            PropertyNameLength = null;
            Level = null;
            Field = 0;
        }

        internal void FromMemberProperty(MemberProperty mp)
        {
            SetAllNull();

            if (mp.Name != null) Name = mp.Name.Value;
            if (mp.ShowCell != null) ShowCell = mp.ShowCell.Value;
            if (mp.ShowTip != null) ShowTip = mp.ShowTip.Value;
            if (mp.ShowAsCaption != null) ShowAsCaption = mp.ShowAsCaption.Value;
            if (mp.NameLength != null) NameLength = mp.NameLength.Value;
            if (mp.PropertyNamePosition != null) PropertyNamePosition = mp.PropertyNamePosition.Value;
            if (mp.PropertyNameLength != null) PropertyNameLength = mp.PropertyNameLength.Value;
            if (mp.Level != null) Level = mp.Level.Value;
            if (mp.Field != null) Field = mp.Field.Value;
        }

        internal MemberProperty ToMemberProperty()
        {
            var mp = new MemberProperty();
            if ((Name != null) && (Name.Length > 0)) mp.Name = Name;
            if (ShowCell) mp.ShowCell = ShowCell;
            if (ShowTip) mp.ShowTip = ShowTip;
            if (ShowAsCaption) mp.ShowAsCaption = ShowAsCaption;
            if (NameLength != null) mp.NameLength = NameLength.Value;
            if (PropertyNamePosition != null) mp.PropertyNamePosition = PropertyNamePosition.Value;
            if (PropertyNameLength != null) mp.PropertyNameLength = PropertyNameLength.Value;
            if (Level != null) mp.Level = Level.Value;
            mp.Field = Field;

            return mp;
        }

        internal SLMemberProperty Clone()
        {
            var mp = new SLMemberProperty();
            mp.Name = Name;
            mp.ShowCell = ShowCell;
            mp.ShowTip = ShowTip;
            mp.ShowAsCaption = ShowAsCaption;
            mp.NameLength = NameLength;
            mp.PropertyNamePosition = PropertyNamePosition;
            mp.PropertyNameLength = PropertyNameLength;
            mp.Level = Level;
            mp.Field = Field;

            return mp;
        }
    }
}