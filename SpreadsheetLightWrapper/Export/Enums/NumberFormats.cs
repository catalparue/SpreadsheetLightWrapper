﻿using System;

namespace SpreadsheetLightWrapper.Export.Enums
{
    /// ===========================================================================================
    /// <summary>
    ///     This Enum assists the user in choosing excel formats from a list of predefined formats.
    /// </summary>
    /// ===========================================================================================
    public enum NumberFormats
    {
        /* >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
         *  Misc
         * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>*/

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "User-Defined" - If the user wants to use a custom format, then there
        ///     be an a valid excel format entered in the UserDefinedNumberFormat field.
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("")] UserDefined = 0,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "General"  36000.1234 -45 - number with no commas; negative values are preceded by a
        ///     minus sign, numbers are not rounded and text as text.
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("@")] General = 1,

        /* >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
         *  Decimals
         * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>*/

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Decimal-0" 42,050 - 0 decimal places with commas
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)")] Decimal0 = 2,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Decimal-1" 42,050.2 - 1 decimal place with commas
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("_(* #,##0.0_);_(* (#,##0.0);_(* \"-\"??_);_(@_)")] Decimal1 = 3,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Decimal-2" 42,050.23 - 2 decimal places with commas
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)")] Decimal2 = 4,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Decimal-3" 42,050.236 - 3 decimal places with commas
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("_(* #,##0.000_);_(* (#,##0.000);_(* \"-\"??_);_(@_)")] Decimal3 = 5,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Decimal-4"  42,050.2356 - 4 decimal places with commas
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("_(* #,##0.0000_);_(* (#,##0.0000);_(* \"-\"??_);_(@_)")] Decimal4 = 6,

        /* >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
         *  Percent
         * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>*/

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Percent-0"  0% - 0 decimal places
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("0%")] Percent0 = 7,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Percent-1"  0.0% - 1 decimal place
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("0.0%")] Percent1 = 8,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Percent-2"  0.00% - 2 decimal places
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("0.00%")] Percent2 = 9,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Percent-3"  0.000% - 3 decimal places
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("0.000%")] Percent3 = 10,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Percent-4"  0.0000% - 4 decimal places
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("0.0000%")] Percent4 = 11,

        /* >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
         *  Currency
         * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>*/

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Currency-0-Black" $36,000 ($45) "black" - 0 decimal places; negative values are
        ///     "Black" in parentheses
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("$#,##0_);($#,##0)")] Currency0Black = 12,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Currency-0-Red" $36,000 ($45) "red" - 0 decimal places; negative values are
        ///     "Red" in parentheses
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("$#,##0_);[Red]($#,##0)")] Currency0Red = 13,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Currency-2-Black" $36,000.00 ($45.00) "black" - 2 decimal places; negative values are
        ///     "Black" in parentheses
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("$#,##0.00_);($#,##0.00)")] Currency2Black = 14,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Currency-2-Red" 36,000.00 ($45.00) "red" - 2 decimal places; negative values are
        ///     "Red" in parentheses
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("$#,##0.00_);[Red]($#,##0.00)")] Currency2Red = 15,


        /* >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
         *  Accounting
         * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>*/

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Accounting-0-Black" 0 decimal places; negative values are "Black" in parentheses
        ///     $		36,000  $		(45) "Black"
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)")] Accounting0Black = 16,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Accounting-0-Red" 0 decimal places; negative values are "Red" in parentheses
        ///     $		36,000  $		(45) "Red"
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("_($* #,##0_);[Red]_($* (#,##0);_($* \"-\"??_);_(@_)")] Accounting0Red = 17,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Accounting-2-Black" 2 decimal places; negative values are "Black" in parentheses
        ///     $		36,000.00  $		(45.00) "Black"
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"_);_(@_)")] Accounting2Black = 18,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Accounting-2-Red" 2 decimal places; negative values are "Red" in parentheses
        ///     $		36,000.00  $		(45.00) "Red"
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("_($* #,##0.00_);[Red]_($* (#,##0.00);_($* \"-\"??_);_(@_)")] Accounting2Red = 19,


        /* >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
         *  Short Dates
         * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>*/

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Date-Short-1"   9/5/2016
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("m/d/yyyy")] DateShort1 = 20,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Date-Short-2"   09/05/2016
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("mm/dd/yyyy")] DateShort2 = 21,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Date-Short-3"   09/5
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("mm/d")] DateShort3 = 22,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Date-Short-4"   5-Sep
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("d-mmm")] DateShort4 = 23,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Date-Short-5"   5-Sep-16
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("d-mmm-yy")] DateShort5 = 24,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Date-Short-6"   05-Sep-16
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("dd-mmm-yy")] DateShort6 = 25,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Date-Short-7"   Sep-5
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("mmm-d")] DateShort7 = 26,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Date-Short-8"   September-5
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("mmmm-d")] DateShort8 = 27,


        /* >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
         *  General Dates
         * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>*/

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Date-General-1"   2016/9/5
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("yyyy/m/d")] DateGeneral1 = 28,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Date-General-2"   2016/09/05
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("yyyy/mm/dd")] DateGeneral2 = 29,

        /* >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
         *  Long Dates
         * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>*/

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Date-Long-1"   Sep 9, 2016
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("mmm d, yyyy")] DateLong1 = 30,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Date-Long-2"   Fri, Sep 9, 2016
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("ddd, mmm d, yyyy")] DateLong2 = 31,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Date-Long-3"   Friday, Sep 9, 2016
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("dddd, mmm d, yyyy")] DateLong3 = 32,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Date-Long-4"   Friday, September 9, 2016
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("dddd, mmmm d, yyyy")] DateLong4 = 33,

        /* >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
         *  Times
         * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>*/

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Time-1-12"  9:30 PM - 12 hour clock
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("h:mm AM/PM")] Time112 = 34,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Time-1-24"  21:30 - 24 hour clock
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("h:mm")] Time124 = 35,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Time-2-12"  09:30 PM - 12 hour clock
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("hh:mm AM/PM")] Time212 = 36,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Time-2-24"  21:30 - 24 hour clock
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("hh:mm")] Time224 = 37,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Time-3-12"  09:30:45 PM - 12 hour clock
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("hh:mm:ss AM/PM")] Time312 = 38,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "Time-3-24"  21:30:45 - 24 hour clock
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("hh:mm:ss")] Time324 = 39,

        /* >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
         *  Timestamps
         * >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>*/

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "TimeStamp-1-12"  2016/09/05 02:42:15 PM - 12 hour timestamp
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("yyyy/mm/dd hh:mm:ss AM/PM")] TimeStamp112 = 40,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "TimeStamp-1-24"  2016/09/05 14:42:15 - 24 hour timestamp
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("yyyy/mm/dd hh:mm:ss")] TimeStamp124 = 41,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "TimeStamp-2-12"  2016/09/05 02:42:15.544 PM - 12 hour timestamp with Milliseconds
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("yyyy/mm/dd hh:mm:ss.000 AM/PM")] TimeStamp212 = 42,

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     "TimeStamp-2-24"  2016/09/05 14:42:15.544 - 24 hour timestamp with Milliseconds
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [FormatString("yyyy/mm/dd hh:mm:ss.000")] TimeStamp224 = 43
    }

    /// ===========================================================================================
    /// <summary>
    ///     Utility:  Creates an string attribute annotation for the Enum items in "NumberFormats"
    /// </summary>
    /// ===========================================================================================
    public class FormatString : Attribute
    {
        public FormatString(string value)
        {
            Value = value;
        }

        public string Value { get; set; }
    }
}