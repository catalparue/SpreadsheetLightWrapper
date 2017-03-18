using DocumentFormat.OpenXml.Spreadsheet;

namespace Ups.Toolkit.SpreadsheetLight.Core.table
{
    /// <summary>
    ///     Specifies the built-in table style type.
    /// </summary>
    public enum SLTableStyleTypeValues
    {
        /// <summary>
        ///     Table Style Light 1
        /// </summary>
        Light1 = 0,

        /// <summary>
        ///     Table Style Light 2
        /// </summary>
        Light2,

        /// <summary>
        ///     Table Style Light 3
        /// </summary>
        Light3,

        /// <summary>
        ///     Table Style Light 4
        /// </summary>
        Light4,

        /// <summary>
        ///     Table Style Light 5
        /// </summary>
        Light5,

        /// <summary>
        ///     Table Style Light 6
        /// </summary>
        Light6,

        /// <summary>
        ///     Table Style Light 7
        /// </summary>
        Light7,

        /// <summary>
        ///     Table Style Light 8
        /// </summary>
        Light8,

        /// <summary>
        ///     Table Style Light 9
        /// </summary>
        Light9,

        /// <summary>
        ///     Table Style Light 10
        /// </summary>
        Light10,

        /// <summary>
        ///     Table Style Light 11
        /// </summary>
        Light11,

        /// <summary>
        ///     Table Style Light 12
        /// </summary>
        Light12,

        /// <summary>
        ///     Table Style Light 13
        /// </summary>
        Light13,

        /// <summary>
        ///     Table Style Light 14
        /// </summary>
        Light14,

        /// <summary>
        ///     Table Style Light 15
        /// </summary>
        Light15,

        /// <summary>
        ///     Table Style Light 16
        /// </summary>
        Light16,

        /// <summary>
        ///     Table Style Light 17
        /// </summary>
        Light17,

        /// <summary>
        ///     Table Style Light 18
        /// </summary>
        Light18,

        /// <summary>
        ///     Table Style Light 19
        /// </summary>
        Light19,

        /// <summary>
        ///     Table Style Light 20
        /// </summary>
        Light20,

        /// <summary>
        ///     Table Style Light 21
        /// </summary>
        Light21,

        /// <summary>
        ///     Table Style Medium 1
        /// </summary>
        Medium1,

        /// <summary>
        ///     Table Style Medium 2
        /// </summary>
        Medium2,

        /// <summary>
        ///     Table Style Medium 3
        /// </summary>
        Medium3,

        /// <summary>
        ///     Table Style Medium 4
        /// </summary>
        Medium4,

        /// <summary>
        ///     Table Style Medium 5
        /// </summary>
        Medium5,

        /// <summary>
        ///     Table Style Medium 6
        /// </summary>
        Medium6,

        /// <summary>
        ///     Table Style Medium 7
        /// </summary>
        Medium7,

        /// <summary>
        ///     Table Style Medium 8
        /// </summary>
        Medium8,

        /// <summary>
        ///     Table Style Medium 9
        /// </summary>
        Medium9,

        /// <summary>
        ///     Table Style Medium 10
        /// </summary>
        Medium10,

        /// <summary>
        ///     Table Style Medium 11
        /// </summary>
        Medium11,

        /// <summary>
        ///     Table Style Medium 12
        /// </summary>
        Medium12,

        /// <summary>
        ///     Table Style Medium 13
        /// </summary>
        Medium13,

        /// <summary>
        ///     Table Style Medium 14
        /// </summary>
        Medium14,

        /// <summary>
        ///     Table Style Medium 15
        /// </summary>
        Medium15,

        /// <summary>
        ///     Table Style Medium 16
        /// </summary>
        Medium16,

        /// <summary>
        ///     Table Style Medium 17
        /// </summary>
        Medium17,

        /// <summary>
        ///     Table Style Medium 18
        /// </summary>
        Medium18,

        /// <summary>
        ///     Table Style Medium 19
        /// </summary>
        Medium19,

        /// <summary>
        ///     Table Style Medium 20
        /// </summary>
        Medium20,

        /// <summary>
        ///     Table Style Medium 21
        /// </summary>
        Medium21,

        /// <summary>
        ///     Table Style Medium 22
        /// </summary>
        Medium22,

        /// <summary>
        ///     Table Style Medium 23
        /// </summary>
        Medium23,

        /// <summary>
        ///     Table Style Medium 24
        /// </summary>
        Medium24,

        /// <summary>
        ///     Table Style Medium 25
        /// </summary>
        Medium25,

        /// <summary>
        ///     Table Style Medium 26
        /// </summary>
        Medium26,

        /// <summary>
        ///     Table Style Medium 27
        /// </summary>
        Medium27,

        /// <summary>
        ///     Table Style Medium 28
        /// </summary>
        Medium28,

        /// <summary>
        ///     Table Style Dark 1
        /// </summary>
        Dark1,

        /// <summary>
        ///     Table Style Dark 2
        /// </summary>
        Dark2,

        /// <summary>
        ///     Table Style Dark 3
        /// </summary>
        Dark3,

        /// <summary>
        ///     Table Style Dark 4
        /// </summary>
        Dark4,

        /// <summary>
        ///     Table Style Dark 5
        /// </summary>
        Dark5,

        /// <summary>
        ///     Table Style Dark 6
        /// </summary>
        Dark6,

        /// <summary>
        ///     Table Style Dark 7
        /// </summary>
        Dark7,

        /// <summary>
        ///     Table Style Dark 8
        /// </summary>
        Dark8,

        /// <summary>
        ///     Table Style Dark 9
        /// </summary>
        Dark9,

        /// <summary>
        ///     Table Style Dark 10
        /// </summary>
        Dark10,

        /// <summary>
        ///     Table Style Dark 11
        /// </summary>
        Dark11
    }

    internal class SLTableStyleInfo
    {
        internal SLTableStyleInfo()
        {
            SetAllNull();
        }

        internal string Name { get; set; }
        internal bool? ShowFirstColumn { get; set; }
        internal bool? ShowLastColumn { get; set; }
        internal bool? ShowRowStripes { get; set; }
        internal bool? ShowColumnStripes { get; set; }

        private void SetAllNull()
        {
            Name = null;
            ShowFirstColumn = null;
            ShowLastColumn = null;
            ShowRowStripes = null;
            ShowColumnStripes = null;
        }

        internal void FromTableStyleInfo(TableStyleInfo tsi)
        {
            SetAllNull();

            if (tsi.Name != null) Name = tsi.Name.Value;
            if (tsi.ShowFirstColumn != null) ShowFirstColumn = tsi.ShowFirstColumn.Value;
            if (tsi.ShowLastColumn != null) ShowLastColumn = tsi.ShowLastColumn.Value;
            if (tsi.ShowRowStripes != null) ShowRowStripes = tsi.ShowRowStripes.Value;
            if (tsi.ShowColumnStripes != null) ShowColumnStripes = tsi.ShowColumnStripes.Value;
        }

        internal TableStyleInfo ToTableStyleInfo()
        {
            var tsi = new TableStyleInfo();
            if (Name != null) tsi.Name = Name;
            if (ShowFirstColumn != null) tsi.ShowFirstColumn = ShowFirstColumn.Value;
            if (ShowLastColumn != null) tsi.ShowLastColumn = ShowLastColumn.Value;
            if (ShowRowStripes != null) tsi.ShowRowStripes = ShowRowStripes.Value;
            if (ShowColumnStripes != null) tsi.ShowColumnStripes = ShowColumnStripes.Value;

            return tsi;
        }

        internal void SetTableStyle(SLTableStyleTypeValues tstyle)
        {
            switch (tstyle)
            {
                case SLTableStyleTypeValues.Light1:
                    Name = "TableStyleLight1";
                    break;
                case SLTableStyleTypeValues.Light2:
                    Name = "TableStyleLight2";
                    break;
                case SLTableStyleTypeValues.Light3:
                    Name = "TableStyleLight3";
                    break;
                case SLTableStyleTypeValues.Light4:
                    Name = "TableStyleLight4";
                    break;
                case SLTableStyleTypeValues.Light5:
                    Name = "TableStyleLight5";
                    break;
                case SLTableStyleTypeValues.Light6:
                    Name = "TableStyleLight6";
                    break;
                case SLTableStyleTypeValues.Light7:
                    Name = "TableStyleLight7";
                    break;
                case SLTableStyleTypeValues.Light8:
                    Name = "TableStyleLight8";
                    break;
                case SLTableStyleTypeValues.Light9:
                    Name = "TableStyleLight9";
                    break;
                case SLTableStyleTypeValues.Light10:
                    Name = "TableStyleLight10";
                    break;
                case SLTableStyleTypeValues.Light11:
                    Name = "TableStyleLight11";
                    break;
                case SLTableStyleTypeValues.Light12:
                    Name = "TableStyleLight12";
                    break;
                case SLTableStyleTypeValues.Light13:
                    Name = "TableStyleLight13";
                    break;
                case SLTableStyleTypeValues.Light14:
                    Name = "TableStyleLight14";
                    break;
                case SLTableStyleTypeValues.Light15:
                    Name = "TableStyleLight15";
                    break;
                case SLTableStyleTypeValues.Light16:
                    Name = "TableStyleLight16";
                    break;
                case SLTableStyleTypeValues.Light17:
                    Name = "TableStyleLight17";
                    break;
                case SLTableStyleTypeValues.Light18:
                    Name = "TableStyleLight18";
                    break;
                case SLTableStyleTypeValues.Light19:
                    Name = "TableStyleLight19";
                    break;
                case SLTableStyleTypeValues.Light20:
                    Name = "TableStyleLight20";
                    break;
                case SLTableStyleTypeValues.Light21:
                    Name = "TableStyleLight21";
                    break;
                case SLTableStyleTypeValues.Medium1:
                    Name = "TableStyleMedium1";
                    break;
                case SLTableStyleTypeValues.Medium2:
                    Name = "TableStyleMedium2";
                    break;
                case SLTableStyleTypeValues.Medium3:
                    Name = "TableStyleMedium3";
                    break;
                case SLTableStyleTypeValues.Medium4:
                    Name = "TableStyleMedium4";
                    break;
                case SLTableStyleTypeValues.Medium5:
                    Name = "TableStyleMedium5";
                    break;
                case SLTableStyleTypeValues.Medium6:
                    Name = "TableStyleMedium6";
                    break;
                case SLTableStyleTypeValues.Medium7:
                    Name = "TableStyleMedium7";
                    break;
                case SLTableStyleTypeValues.Medium8:
                    Name = "TableStyleMedium8";
                    break;
                case SLTableStyleTypeValues.Medium9:
                    Name = "TableStyleMedium9";
                    break;
                case SLTableStyleTypeValues.Medium10:
                    Name = "TableStyleMedium10";
                    break;
                case SLTableStyleTypeValues.Medium11:
                    Name = "TableStyleMedium11";
                    break;
                case SLTableStyleTypeValues.Medium12:
                    Name = "TableStyleMedium12";
                    break;
                case SLTableStyleTypeValues.Medium13:
                    Name = "TableStyleMedium13";
                    break;
                case SLTableStyleTypeValues.Medium14:
                    Name = "TableStyleMedium14";
                    break;
                case SLTableStyleTypeValues.Medium15:
                    Name = "TableStyleMedium15";
                    break;
                case SLTableStyleTypeValues.Medium16:
                    Name = "TableStyleMedium16";
                    break;
                case SLTableStyleTypeValues.Medium17:
                    Name = "TableStyleMedium17";
                    break;
                case SLTableStyleTypeValues.Medium18:
                    Name = "TableStyleMedium18";
                    break;
                case SLTableStyleTypeValues.Medium19:
                    Name = "TableStyleMedium19";
                    break;
                case SLTableStyleTypeValues.Medium20:
                    Name = "TableStyleMedium20";
                    break;
                case SLTableStyleTypeValues.Medium21:
                    Name = "TableStyleMedium21";
                    break;
                case SLTableStyleTypeValues.Medium22:
                    Name = "TableStyleMedium22";
                    break;
                case SLTableStyleTypeValues.Medium23:
                    Name = "TableStyleMedium23";
                    break;
                case SLTableStyleTypeValues.Medium24:
                    Name = "TableStyleMedium24";
                    break;
                case SLTableStyleTypeValues.Medium25:
                    Name = "TableStyleMedium25";
                    break;
                case SLTableStyleTypeValues.Medium26:
                    Name = "TableStyleMedium26";
                    break;
                case SLTableStyleTypeValues.Medium27:
                    Name = "TableStyleMedium27";
                    break;
                case SLTableStyleTypeValues.Medium28:
                    Name = "TableStyleMedium28";
                    break;
                case SLTableStyleTypeValues.Dark1:
                    Name = "TableStyleDark1";
                    break;
                case SLTableStyleTypeValues.Dark2:
                    Name = "TableStyleDark2";
                    break;
                case SLTableStyleTypeValues.Dark3:
                    Name = "TableStyleDark3";
                    break;
                case SLTableStyleTypeValues.Dark4:
                    Name = "TableStyleDark4";
                    break;
                case SLTableStyleTypeValues.Dark5:
                    Name = "TableStyleDark5";
                    break;
                case SLTableStyleTypeValues.Dark6:
                    Name = "TableStyleDark6";
                    break;
                case SLTableStyleTypeValues.Dark7:
                    Name = "TableStyleDark7";
                    break;
                case SLTableStyleTypeValues.Dark8:
                    Name = "TableStyleDark8";
                    break;
                case SLTableStyleTypeValues.Dark9:
                    Name = "TableStyleDark9";
                    break;
                case SLTableStyleTypeValues.Dark10:
                    Name = "TableStyleDark10";
                    break;
                case SLTableStyleTypeValues.Dark11:
                    Name = "TableStyleDark11";
                    break;
            }
        }

        internal SLTableStyleInfo Clone()
        {
            var tsi = new SLTableStyleInfo();
            tsi.Name = Name;
            tsi.ShowFirstColumn = ShowFirstColumn;
            tsi.ShowLastColumn = ShowLastColumn;
            tsi.ShowRowStripes = ShowRowStripes;
            tsi.ShowColumnStripes = ShowColumnStripes;

            return tsi;
        }
    }
}