using System;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLightWrapper.Export.Enums;

namespace SpreadsheetLightWrapper.Export.Models
{
    /// ===========================================================================================
    /// <summary>
    ///     Model intended for the setup of User-Defined field properties allowing custom field
    ///     names, formatting, visibility and order.
    /// </summary>
    /// ===========================================================================================
    public class Column
    {
        #region Properties

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     This the Bound column name
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        public string BoundColumnName { get; set; }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Enter a Custom Column name that will replace the Bound Field name
        ///     or leave blank if you want the bound name
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        public string UserDefinedColumnName { get; set; }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Select a NumberFormats Enum value for a Data Formatting.
        ///     <para />
        ///     ------------------------------
        ///     <para />
        ///     "UserDefined" - If the user wants to use a custom format, then there must
        ///     be a valid excel format entered in the "UserDefinedNumberFormat" field.
        ///     <para />
        ///     ------------------------------
        ///     <para />
        ///     "General" 36000.1234 -45 number with no commas; negative values are preceded by a
        ///     minus sign and text as text.
        ///     <para />
        ///     ------------------------------
        ///     <para />
        ///     "Decimal0" 42,050 0 decimal places with commas
        ///     || "Decimal1" 42,050.2 1 decimal place with commas
        ///     || "Decimal2" 42,050.23 2 decimal places with commas
        ///     || "Decimal3" 42,050.233 3 decimal places with commas
        ///     || "Decimal-4" 42,050.2334 4 decimal places with commas
        ///     <para />
        ///     ------------------------------
        ///     <para />
        ///     Input values must be decimal values between 0-1 ->
        ///     "Percent0" 89% 0 decimal places.
        ///     || "Percent1" 89.3% 1 decimal place
        ///     || "Percent2" 89.25% 2 decimal places
        ///     || "Percent3" 89.251% 3 decimal places
        ///     || "Percent4" 89.2515% 4 decimal places
        ///     <para />
        ///     ------------------------------
        ///     <para />
        ///     "Currency0Black" $36,000  ($45) "black" 0 decimal places; negative values are "Black" in parentheses
        ///     || "Currency0Red" $36,000 ($45) "red" 0 decimal places; negative values are "Red" in parentheses
        ///     || "Currency2Black" $36,000.12  ($45.00) "black" 2 decimal places; negative values are "Black" in parentheses
        ///     || "Currency2Red" $36,000.12  ($45.00) "red" 2 decimal places; negative values are "Red" in parentheses
        ///     <para />
        ///     ------------------------------
        ///     <para />
        ///     "Accounting0Black" $		36,000  $		(45)
        ///     "Black" 0 decimal places; negative values are "Black" in parentheses
        ///     Same as the Accounting format with the $ on the opposite side of the field
        ///     from the numbers
        ///     || "Accounting0Red" $		36,000  $		(45)
        ///     "Red" 0 decimal places; negative values are "Red" in parentheses
        ///     Same as the Accounting format with the $ on the opposite side of the field
        ///     from the numbers
        ///     || "Accounting2Black" $		36,000.45  $	(45.30) "Black"
        ///     2 decimal places; negative values are "Black" in parentheses
        ///     Same as the Accounting format with the $ on the opposite side of the field
        ///     from the numbers
        ///     || "Accounting2Red" $		36,000.45  $	(45.30) "Black"
        ///     2 decimal places; negative values are "Red" in parentheses
        ///     Same as the Accounting format with the $ on the opposite side of the field
        ///     from the numbers
        ///     <para />
        ///     ------------------------------
        ///     <para />
        ///     "DateShort1" 9/5/2016
        ///     || "DateShort2" 09/05/2016
        ///     || "DateShort3" 09/5
        ///     || "DateShort4" 5-Sep
        ///     || "DateShort5" 5-Sep-16
        ///     || "DateShort6" 05-Sep-16
        ///     || "DateShort7" Sep-5
        ///     || "DateShort8" September-5
        ///     || "DateGeneral1" 2016/9/5
        ///     || "DateGeneral2" 2016/09/05
        ///     || "DateLong1" Sep 9, 2016
        ///     || "DateLong2" Fri, Sep 9, 2016
        ///     || "DateLong3" Friday, Sep 9, 2016
        ///     || "DateLong4" Friday, September 9, 2016
        ///     <para />
        ///     ------------------------------
        ///     <para />
        ///     "Time112" 9:30 PM - 12 hour clock
        ///     || "Time124" 21:30 - 24 hour clock
        ///     || "Time212" 09:30 PM - 12 hour clock
        ///     || "Time224" 21:30 - 24 hour clock
        ///     || "Time312" 09:30:45 PM - 12 hour clock
        ///     || "Time324" 21:30:45 - 24 hour clock
        ///     <para />
        ///     ------------------------------
        ///     <para />
        ///     "TimeStamp112" 2016/09/05 02:42:15 PM - 12 hour timestamp
        ///     || "TimeStamp124" 2016/09/05 14:42:15 - 24 hour timestamp
        ///     || "TimeStamp212" 2016/09/05 02:42:15.544 PM - 12 hour timestamp with Milliseconds
        ///     || "TimeStamp224" 2016/09/05 14:42:15.544 - 24 hour timestamp with Milliseconds
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        public NumberFormats NumberFormat { get; set; }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Enter a custom user-defined excel format
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        public string UserDefinedNumberFormat { get; set; }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Select an HorizontalAlignmentValues Enum value of "Left", "Center" or "Right" for
        ///     horizontal alignment.
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        public HorizontalAlignmentValues HorizontalAlignment { get; set; }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Enter a boolean value of true or false to show or hide the field.
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        public bool ShowField { get; set; }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Enter an integer value of 1, 2, 3, 4, etc. to set the field display order.
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        public int? FieldOrder { get; set; }

        #endregion Properties

        #region Constructors

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Overload 1: Constructor
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        public Column()
        {
        }


        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Overload 2: Constructor, Initialize all fields
        /// </summary>
        /// <param name="boundColumnName">string</param>
        /// <param name="userDefinedColumnName">string</param>
        /// <param name="numberFormat"></param>
        /// <param name="horizontalAlignment"></param>
        /// <param name="showField">bool</param>
        /// <param name="fieldOrder">int?</param>
        /// <param name="userDefinedNumberFormat">string</param>
        /// -----------------------------------------------------------------------------------------------
        public Column(
            string boundColumnName,
            string userDefinedColumnName,
            NumberFormats numberFormat = NumberFormats.General,
            HorizontalAlignmentValues horizontalAlignment = HorizontalAlignmentValues.Center,
            bool showField = true,
            int? fieldOrder = null,
            string userDefinedNumberFormat = null)
        {
            try
            {
                BoundColumnName = boundColumnName;
                UserDefinedColumnName = userDefinedColumnName;
                NumberFormat = numberFormat;
                HorizontalAlignment = horizontalAlignment;
                ShowField = showField;
                FieldOrder = fieldOrder;
                UserDefinedNumberFormat = userDefinedNumberFormat;
            }
            catch (Exception ex)
            {
                //WebLogger.LogException(
                //    new Exception(
                //        "Ups.Toolkit.SpreadsheetLight.Export.Models.Column.Contructor:Overload 2 -> " +
                //        ex.Message, ex),
                //    new Dictionary<string, string> { { "Column", "Constructor:Overload 2" } });
            }
        }

        #endregion Constructors
    }
}
