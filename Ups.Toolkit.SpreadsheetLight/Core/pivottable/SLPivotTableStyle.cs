using DocumentFormat.OpenXml.Spreadsheet;
using Ups.Toolkit.SpreadsheetLight.Core.misc;

namespace Ups.Toolkit.SpreadsheetLight.Core.pivottable
{
    public enum SLPivotTableStyleTypeValues
    {
        /// <summary>
        ///     Pivot Style Light 1
        /// </summary>
        Light1 = 0,

        /// <summary>
        ///     Pivot Style Light 2
        /// </summary>
        Light2,

        /// <summary>
        ///     Pivot Style Light 3
        /// </summary>
        Light3,

        /// <summary>
        ///     Pivot Style Light 4
        /// </summary>
        Light4,

        /// <summary>
        ///     Pivot Style Light 5
        /// </summary>
        Light5,

        /// <summary>
        ///     Pivot Style Light 6
        /// </summary>
        Light6,

        /// <summary>
        ///     Pivot Style Light 7
        /// </summary>
        Light7,

        /// <summary>
        ///     Pivot Style Light 8
        /// </summary>
        Light8,

        /// <summary>
        ///     Pivot Style Light 9
        /// </summary>
        Light9,

        /// <summary>
        ///     Pivot Style Light 10
        /// </summary>
        Light10,

        /// <summary>
        ///     Pivot Style Light 11
        /// </summary>
        Light11,

        /// <summary>
        ///     Pivot Style Light 12
        /// </summary>
        Light12,

        /// <summary>
        ///     Pivot Style Light 13
        /// </summary>
        Light13,

        /// <summary>
        ///     Pivot Style Light 14
        /// </summary>
        Light14,

        /// <summary>
        ///     Pivot Style Light 15
        /// </summary>
        Light15,

        /// <summary>
        ///     Pivot Style Light 16
        /// </summary>
        Light16,

        /// <summary>
        ///     Pivot Style Light 17
        /// </summary>
        Light17,

        /// <summary>
        ///     Pivot Style Light 18
        /// </summary>
        Light18,

        /// <summary>
        ///     Pivot Style Light 19
        /// </summary>
        Light19,

        /// <summary>
        ///     Pivot Style Light 20
        /// </summary>
        Light20,

        /// <summary>
        ///     Pivot Style Light 21
        /// </summary>
        Light21,

        /// <summary>
        ///     Pivot Style Light 22
        /// </summary>
        Light22,

        /// <summary>
        ///     Pivot Style Light 23
        /// </summary>
        Light23,

        /// <summary>
        ///     Pivot Style Light 24
        /// </summary>
        Light24,

        /// <summary>
        ///     Pivot Style Light 25
        /// </summary>
        Light25,

        /// <summary>
        ///     Pivot Style Light 26
        /// </summary>
        Light26,

        /// <summary>
        ///     Pivot Style Light 27
        /// </summary>
        Light27,

        /// <summary>
        ///     Pivot Style Light 28
        /// </summary>
        Light28,

        /// <summary>
        ///     Pivot Style Medium 1
        /// </summary>
        Medium1,

        /// <summary>
        ///     Pivot Style Medium 2
        /// </summary>
        Medium2,

        /// <summary>
        ///     Pivot Style Medium 3
        /// </summary>
        Medium3,

        /// <summary>
        ///     Pivot Style Medium 4
        /// </summary>
        Medium4,

        /// <summary>
        ///     Pivot Style Medium 5
        /// </summary>
        Medium5,

        /// <summary>
        ///     Pivot Style Medium 6
        /// </summary>
        Medium6,

        /// <summary>
        ///     Pivot Style Medium 7
        /// </summary>
        Medium7,

        /// <summary>
        ///     Pivot Style Medium 8
        /// </summary>
        Medium8,

        /// <summary>
        ///     Pivot Style Medium 9
        /// </summary>
        Medium9,

        /// <summary>
        ///     Pivot Style Medium 10
        /// </summary>
        Medium10,

        /// <summary>
        ///     Pivot Style Medium 11
        /// </summary>
        Medium11,

        /// <summary>
        ///     Pivot Style Medium 12
        /// </summary>
        Medium12,

        /// <summary>
        ///     Pivot Style Medium 13
        /// </summary>
        Medium13,

        /// <summary>
        ///     Pivot Style Medium 14
        /// </summary>
        Medium14,

        /// <summary>
        ///     Pivot Style Medium 15
        /// </summary>
        Medium15,

        /// <summary>
        ///     Pivot Style Medium 16
        /// </summary>
        Medium16,

        /// <summary>
        ///     Pivot Style Medium 17
        /// </summary>
        Medium17,

        /// <summary>
        ///     Pivot Style Medium 18
        /// </summary>
        Medium18,

        /// <summary>
        ///     Pivot Style Medium 19
        /// </summary>
        Medium19,

        /// <summary>
        ///     Pivot Style Medium 20
        /// </summary>
        Medium20,

        /// <summary>
        ///     Pivot Style Medium 21
        /// </summary>
        Medium21,

        /// <summary>
        ///     Pivot Style Medium 22
        /// </summary>
        Medium22,

        /// <summary>
        ///     Pivot Style Medium 23
        /// </summary>
        Medium23,

        /// <summary>
        ///     Pivot Style Medium 24
        /// </summary>
        Medium24,

        /// <summary>
        ///     Pivot Style Medium 25
        /// </summary>
        Medium25,

        /// <summary>
        ///     Pivot Style Medium 26
        /// </summary>
        Medium26,

        /// <summary>
        ///     Pivot Style Medium 27
        /// </summary>
        Medium27,

        /// <summary>
        ///     Pivot Style Medium 28
        /// </summary>
        Medium28,

        /// <summary>
        ///     Pivot Style Dark 1
        /// </summary>
        Dark1,

        /// <summary>
        ///     Pivot Style Dark 2
        /// </summary>
        Dark2,

        /// <summary>
        ///     Pivot Style Dark 3
        /// </summary>
        Dark3,

        /// <summary>
        ///     Pivot Style Dark 4
        /// </summary>
        Dark4,

        /// <summary>
        ///     Pivot Style Dark 5
        /// </summary>
        Dark5,

        /// <summary>
        ///     Pivot Style Dark 6
        /// </summary>
        Dark6,

        /// <summary>
        ///     Pivot Style Dark 7
        /// </summary>
        Dark7,

        /// <summary>
        ///     Pivot Style Dark 8
        /// </summary>
        Dark8,

        /// <summary>
        ///     Pivot Style Dark 9
        /// </summary>
        Dark9,

        /// <summary>
        ///     Pivot Style Dark 10
        /// </summary>
        Dark10,

        /// <summary>
        ///     Pivot Style Dark 11
        /// </summary>
        Dark11,

        /// <summary>
        ///     Pivot Style Dark 12
        /// </summary>
        Dark12,

        /// <summary>
        ///     Pivot Style Dark 13
        /// </summary>
        Dark13,

        /// <summary>
        ///     Pivot Style Dark 14
        /// </summary>
        Dark14,

        /// <summary>
        ///     Pivot Style Dark 15
        /// </summary>
        Dark15,

        /// <summary>
        ///     Pivot Style Dark 16
        /// </summary>
        Dark16,

        /// <summary>
        ///     Pivot Style Dark 17
        /// </summary>
        Dark17,

        /// <summary>
        ///     Pivot Style Dark 18
        /// </summary>
        Dark18,

        /// <summary>
        ///     Pivot Style Dark 19
        /// </summary>
        Dark19,

        /// <summary>
        ///     Pivot Style Dark 20
        /// </summary>
        Dark20,

        /// <summary>
        ///     Pivot Style Dark 21
        /// </summary>
        Dark21,

        /// <summary>
        ///     Pivot Style Dark 22
        /// </summary>
        Dark22,

        /// <summary>
        ///     Pivot Style Dark 23
        /// </summary>
        Dark23,

        /// <summary>
        ///     Pivot Style Dark 24
        /// </summary>
        Dark24,

        /// <summary>
        ///     Pivot Style Dark 25
        /// </summary>
        Dark25,

        /// <summary>
        ///     Pivot Style Dark 26
        /// </summary>
        Dark26,

        /// <summary>
        ///     Pivot Style Dark 27
        /// </summary>
        Dark27,

        /// <summary>
        ///     Pivot Style Dark 28
        /// </summary>
        Dark28
    }

    internal class SLPivotTableStyle
    {
        internal SLPivotTableStyle()
        {
            SetAllNull();
        }

        internal string Name { get; set; }
        internal bool ShowRowHeaders { get; set; }
        internal bool ShowColumnHeaders { get; set; }
        internal bool ShowRowStripes { get; set; }
        internal bool ShowColumnStripes { get; set; }
        internal bool? ShowLastColumn { get; set; }

        private void SetAllNull()
        {
            Name = SLConstants.DefaultPivotStyle;
            ShowRowHeaders = true;
            ShowColumnHeaders = true;
            ShowRowStripes = false;
            ShowColumnStripes = false;
            ShowLastColumn = true;
        }

        internal void FromPivotTableStyle(PivotTableStyle pts)
        {
            SetAllNull();

            if (pts.Name != null) Name = pts.Name.Value;
            if (pts.ShowRowHeaders != null) ShowRowHeaders = pts.ShowRowHeaders.Value;
            if (pts.ShowColumnHeaders != null) ShowColumnHeaders = pts.ShowColumnHeaders.Value;
            if (pts.ShowRowStripes != null) ShowRowStripes = pts.ShowRowStripes.Value;
            if (pts.ShowColumnStripes != null) ShowColumnStripes = pts.ShowColumnStripes.Value;
            if (pts.ShowLastColumn != null) ShowLastColumn = pts.ShowLastColumn.Value;
        }

        internal PivotTableStyle ToPivotTableStyle()
        {
            var pts = new PivotTableStyle();
            if ((Name != null) && (Name.Length > 0)) pts.Name = Name;
            pts.ShowRowHeaders = ShowRowHeaders;
            pts.ShowColumnHeaders = ShowColumnHeaders;
            pts.ShowRowStripes = ShowRowStripes;
            pts.ShowColumnStripes = ShowColumnStripes;
            if (ShowLastColumn != null) pts.ShowLastColumn = ShowLastColumn.Value;

            return pts;
        }

        internal void SetPivotTableStyle(SLPivotTableStyleTypeValues pivotstyle)
        {
            switch (pivotstyle)
            {
                case SLPivotTableStyleTypeValues.Light1:
                    Name = "PivotStyleLight1";
                    break;
                case SLPivotTableStyleTypeValues.Light2:
                    Name = "PivotStyleLight2";
                    break;
                case SLPivotTableStyleTypeValues.Light3:
                    Name = "PivotStyleLight3";
                    break;
                case SLPivotTableStyleTypeValues.Light4:
                    Name = "PivotStyleLight4";
                    break;
                case SLPivotTableStyleTypeValues.Light5:
                    Name = "PivotStyleLight5";
                    break;
                case SLPivotTableStyleTypeValues.Light6:
                    Name = "PivotStyleLight6";
                    break;
                case SLPivotTableStyleTypeValues.Light7:
                    Name = "PivotStyleLight7";
                    break;
                case SLPivotTableStyleTypeValues.Light8:
                    Name = "PivotStyleLight8";
                    break;
                case SLPivotTableStyleTypeValues.Light9:
                    Name = "PivotStyleLight9";
                    break;
                case SLPivotTableStyleTypeValues.Light10:
                    Name = "PivotStyleLight10";
                    break;
                case SLPivotTableStyleTypeValues.Light11:
                    Name = "PivotStyleLight11";
                    break;
                case SLPivotTableStyleTypeValues.Light12:
                    Name = "PivotStyleLight12";
                    break;
                case SLPivotTableStyleTypeValues.Light13:
                    Name = "PivotStyleLight13";
                    break;
                case SLPivotTableStyleTypeValues.Light14:
                    Name = "PivotStyleLight14";
                    break;
                case SLPivotTableStyleTypeValues.Light15:
                    Name = "PivotStyleLight15";
                    break;
                case SLPivotTableStyleTypeValues.Light16:
                    Name = "PivotStyleLight16";
                    break;
                case SLPivotTableStyleTypeValues.Light17:
                    Name = "PivotStyleLight17";
                    break;
                case SLPivotTableStyleTypeValues.Light18:
                    Name = "PivotStyleLight18";
                    break;
                case SLPivotTableStyleTypeValues.Light19:
                    Name = "PivotStyleLight19";
                    break;
                case SLPivotTableStyleTypeValues.Light20:
                    Name = "PivotStyleLight20";
                    break;
                case SLPivotTableStyleTypeValues.Light21:
                    Name = "PivotStyleLight21";
                    break;
                case SLPivotTableStyleTypeValues.Light22:
                    Name = "PivotStyleLight22";
                    break;
                case SLPivotTableStyleTypeValues.Light23:
                    Name = "PivotStyleLight23";
                    break;
                case SLPivotTableStyleTypeValues.Light24:
                    Name = "PivotStyleLight24";
                    break;
                case SLPivotTableStyleTypeValues.Light25:
                    Name = "PivotStyleLight25";
                    break;
                case SLPivotTableStyleTypeValues.Light26:
                    Name = "PivotStyleLight26";
                    break;
                case SLPivotTableStyleTypeValues.Light27:
                    Name = "PivotStyleLight27";
                    break;
                case SLPivotTableStyleTypeValues.Light28:
                    Name = "PivotStyleLight28";
                    break;
                case SLPivotTableStyleTypeValues.Medium1:
                    Name = "PivotStyleMedium1";
                    break;
                case SLPivotTableStyleTypeValues.Medium2:
                    Name = "PivotStyleMedium2";
                    break;
                case SLPivotTableStyleTypeValues.Medium3:
                    Name = "PivotStyleMedium3";
                    break;
                case SLPivotTableStyleTypeValues.Medium4:
                    Name = "PivotStyleMedium4";
                    break;
                case SLPivotTableStyleTypeValues.Medium5:
                    Name = "PivotStyleMedium5";
                    break;
                case SLPivotTableStyleTypeValues.Medium6:
                    Name = "PivotStyleMedium6";
                    break;
                case SLPivotTableStyleTypeValues.Medium7:
                    Name = "PivotStyleMedium7";
                    break;
                case SLPivotTableStyleTypeValues.Medium8:
                    Name = "PivotStyleMedium8";
                    break;
                case SLPivotTableStyleTypeValues.Medium9:
                    Name = "PivotStyleMedium9";
                    break;
                case SLPivotTableStyleTypeValues.Medium10:
                    Name = "PivotStyleMedium10";
                    break;
                case SLPivotTableStyleTypeValues.Medium11:
                    Name = "PivotStyleMedium11";
                    break;
                case SLPivotTableStyleTypeValues.Medium12:
                    Name = "PivotStyleMedium12";
                    break;
                case SLPivotTableStyleTypeValues.Medium13:
                    Name = "PivotStyleMedium13";
                    break;
                case SLPivotTableStyleTypeValues.Medium14:
                    Name = "PivotStyleMedium14";
                    break;
                case SLPivotTableStyleTypeValues.Medium15:
                    Name = "PivotStyleMedium15";
                    break;
                case SLPivotTableStyleTypeValues.Medium16:
                    Name = "PivotStyleMedium16";
                    break;
                case SLPivotTableStyleTypeValues.Medium17:
                    Name = "PivotStyleMedium17";
                    break;
                case SLPivotTableStyleTypeValues.Medium18:
                    Name = "PivotStyleMedium18";
                    break;
                case SLPivotTableStyleTypeValues.Medium19:
                    Name = "PivotStyleMedium19";
                    break;
                case SLPivotTableStyleTypeValues.Medium20:
                    Name = "PivotStyleMedium20";
                    break;
                case SLPivotTableStyleTypeValues.Medium21:
                    Name = "PivotStyleMedium21";
                    break;
                case SLPivotTableStyleTypeValues.Medium22:
                    Name = "PivotStyleMedium22";
                    break;
                case SLPivotTableStyleTypeValues.Medium23:
                    Name = "PivotStyleMedium23";
                    break;
                case SLPivotTableStyleTypeValues.Medium24:
                    Name = "PivotStyleMedium24";
                    break;
                case SLPivotTableStyleTypeValues.Medium25:
                    Name = "PivotStyleMedium25";
                    break;
                case SLPivotTableStyleTypeValues.Medium26:
                    Name = "PivotStyleMedium26";
                    break;
                case SLPivotTableStyleTypeValues.Medium27:
                    Name = "PivotStyleMedium27";
                    break;
                case SLPivotTableStyleTypeValues.Medium28:
                    Name = "PivotStyleMedium28";
                    break;
                case SLPivotTableStyleTypeValues.Dark1:
                    Name = "PivotStyleDark1";
                    break;
                case SLPivotTableStyleTypeValues.Dark2:
                    Name = "PivotStyleDark2";
                    break;
                case SLPivotTableStyleTypeValues.Dark3:
                    Name = "PivotStyleDark3";
                    break;
                case SLPivotTableStyleTypeValues.Dark4:
                    Name = "PivotStyleDark4";
                    break;
                case SLPivotTableStyleTypeValues.Dark5:
                    Name = "PivotStyleDark5";
                    break;
                case SLPivotTableStyleTypeValues.Dark6:
                    Name = "PivotStyleDark6";
                    break;
                case SLPivotTableStyleTypeValues.Dark7:
                    Name = "PivotStyleDark7";
                    break;
                case SLPivotTableStyleTypeValues.Dark8:
                    Name = "PivotStyleDark8";
                    break;
                case SLPivotTableStyleTypeValues.Dark9:
                    Name = "PivotStyleDark9";
                    break;
                case SLPivotTableStyleTypeValues.Dark10:
                    Name = "PivotStyleDark10";
                    break;
                case SLPivotTableStyleTypeValues.Dark11:
                    Name = "PivotStyleDark11";
                    break;
                case SLPivotTableStyleTypeValues.Dark12:
                    Name = "PivotStyleDark12";
                    break;
                case SLPivotTableStyleTypeValues.Dark13:
                    Name = "PivotStyleDark13";
                    break;
                case SLPivotTableStyleTypeValues.Dark14:
                    Name = "PivotStyleDark14";
                    break;
                case SLPivotTableStyleTypeValues.Dark15:
                    Name = "PivotStyleDark15";
                    break;
                case SLPivotTableStyleTypeValues.Dark16:
                    Name = "PivotStyleDark16";
                    break;
                case SLPivotTableStyleTypeValues.Dark17:
                    Name = "PivotStyleDark17";
                    break;
                case SLPivotTableStyleTypeValues.Dark18:
                    Name = "PivotStyleDark18";
                    break;
                case SLPivotTableStyleTypeValues.Dark19:
                    Name = "PivotStyleDark19";
                    break;
                case SLPivotTableStyleTypeValues.Dark20:
                    Name = "PivotStyleDark20";
                    break;
                case SLPivotTableStyleTypeValues.Dark21:
                    Name = "PivotStyleDark21";
                    break;
                case SLPivotTableStyleTypeValues.Dark22:
                    Name = "PivotStyleDark22";
                    break;
                case SLPivotTableStyleTypeValues.Dark23:
                    Name = "PivotStyleDark23";
                    break;
                case SLPivotTableStyleTypeValues.Dark24:
                    Name = "PivotStyleDark24";
                    break;
                case SLPivotTableStyleTypeValues.Dark25:
                    Name = "PivotStyleDark25";
                    break;
                case SLPivotTableStyleTypeValues.Dark26:
                    Name = "PivotStyleDark26";
                    break;
                case SLPivotTableStyleTypeValues.Dark27:
                    Name = "PivotStyleDark27";
                    break;
                case SLPivotTableStyleTypeValues.Dark28:
                    Name = "PivotStyleDark28";
                    break;
            }
        }

        internal SLPivotTableStyle Clone()
        {
            var pts = new SLPivotTableStyle();
            pts.Name = Name;
            pts.ShowRowHeaders = ShowRowHeaders;
            pts.ShowColumnHeaders = ShowColumnHeaders;
            pts.ShowRowStripes = ShowRowStripes;
            pts.ShowColumnStripes = ShowColumnStripes;
            pts.ShowLastColumn = ShowLastColumn;

            return pts;
        }
    }
}