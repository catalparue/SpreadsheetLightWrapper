using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLightWrapper.Core.style;
using SpreadsheetLightWrapper.Export.Enums;

namespace SpreadsheetLightWrapper.Export.Models
{
    /// ===========================================================================================
    /// <summary>
    ///     Allows the programmer to set custom properties for multiple datasets in Excel
    /// </summary>
    /// ===========================================================================================
    public class ChildSetting
    {
        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Properties
        /// </summary>
        /// -----------------------------------------------------------------------------------------------

        #region Properties

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Properties
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        public string SheetName { get; set; }
        public bool ShowColumnHeader { get; set; }
        public bool ShowAlternatingRows { get; set; }
        public int ColumnOffset { get; set; }
        public int? ColumnHeaderRowHeight { get; set; }
        public SLStyle EvenRowStyle { get; set; }
        public SLStyle OddRowStyle { get; set; }
        public SLStyle ColumnHeaderStyle { get; set; }
        public List<Column> UserDefinedColumns { get; set; }

        #endregion Properties

        #region Constructors

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Overload 1: Constructor - Base Constructor
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        public ChildSetting()
        {
            try
            {
                SheetName = string.Empty;
                ShowColumnHeader = true;
                ColumnOffset = 0;
                ColumnHeaderRowHeight = null;
                ColumnHeaderStyle = new SLStyle();
                ShowAlternatingRows = false;
                OddRowStyle = new SLStyle();
                EvenRowStyle = new SLStyle();
                UserDefinedColumns = new List<Column>();
            }
            catch (Exception ex)
            {
                //WebLogger.LogException(
                //    new Exception(
                //        "SpreadsheetLightWrapper.Export.Models.ChildSetting.Contructor:Overload 1 -> " +
                //        ex.Message, ex),
                    //new Dictionary<string, string> {{"ChildSetting", "Constructor:Overload 1"}});
            }
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Overload 2: Constructor - Set commonly used properties
        /// </summary>
        /// <param name="columnHeaderStyle">SLStyle</param>
        /// <param name="oddRowStyle">SLStyle</param>
        /// -----------------------------------------------------------------------------------------------
        public ChildSetting(
            SLStyle columnHeaderStyle,
            SLStyle oddRowStyle
        )
        {
            try
            {
                SheetName = string.Empty;
                ShowColumnHeader = true;
                ColumnOffset = 0;
                ColumnHeaderRowHeight = null;
                ColumnHeaderStyle = columnHeaderStyle ?? new SLStyle();
                ShowAlternatingRows = false;
                OddRowStyle = oddRowStyle ?? new SLStyle();
                EvenRowStyle = new SLStyle();
                UserDefinedColumns = new List<Column>();
            }
            catch (Exception ex)
            {
                //WebLogger.LogException(
                //    new Exception(
                //        "SpreadsheetLightWrapper.Export.Models.ChildSetting.Contructor:Overload 2 -> " +
                //        ex.Message, ex),
                //    new Dictionary<string, string> {{"ChildSetting", "Constructor:Overload 2"}});
            }
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Overload 3: Constructor - Set commonly used properties
        /// </summary>
        /// <param name="columnHeaderStyle">SLStyle</param>
        /// <param name="showAlternatingRows">bool?</param>
        /// <param name="oddRowStyle">SLStyle</param>
        /// <param name="evenRowStyle">SLStyle</param>
        /// <param name="userDefinedColumns">List(Column)</param>
        /// -----------------------------------------------------------------------------------------------
        public ChildSetting(
            SLStyle columnHeaderStyle,
            bool? showAlternatingRows,
            SLStyle oddRowStyle,
            SLStyle evenRowStyle,
            List<Column> userDefinedColumns = null
        )
        {
            try
            {
                SheetName = string.Empty;
                ShowColumnHeader = true;
                ColumnOffset = 0;
                ColumnHeaderRowHeight = null;
                ColumnHeaderStyle = columnHeaderStyle ?? new SLStyle();
                ShowAlternatingRows = showAlternatingRows ?? false;
                OddRowStyle = oddRowStyle ?? new SLStyle();
                EvenRowStyle = evenRowStyle ?? new SLStyle();
                UserDefinedColumns = userDefinedColumns ?? new List<Column>();
            }
            catch (Exception ex)
            {
                //WebLogger.LogException(
                //    new Exception(
                //        "SpreadsheetLightWrapper.Export.Models.ChildSetting.Contructor:Overload 3 -> " +
                //        ex.Message, ex),
                //    new Dictionary<string, string> {{"ChildSetting", "Constructor:Overload 3"}});
            }
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Overload 4: Constructor - Set all properties
        /// </summary>
        /// <param name="name">string</param>
        /// <param name="showColumnHeader">bool?</param>
        /// <param name="columnOffset">int?</param>
        /// <param name="columnHeaderRowHeight">int?</param>
        /// <param name="columnHeaderStyle">SLStyle</param>
        /// <param name="evenRowStyle">SLStyle</param>
        /// <param name="showAlternatingRows"></param>
        /// <param name="oddRowStyle">SLStyle</param>
        /// <param name="userDefinedColumns">List(Column)</param>
        /// -----------------------------------------------------------------------------------------------
        public ChildSetting(
            string name,
            bool? showColumnHeader,
            int? columnOffset,
            int? columnHeaderRowHeight,
            SLStyle columnHeaderStyle,
            bool? showAlternatingRows,
            SLStyle oddRowStyle,
            SLStyle evenRowStyle,
            List<Column> userDefinedColumns
        )
        {
            try
            {
                SheetName = name ?? string.Empty;
                ShowColumnHeader = showColumnHeader ?? true;
                ColumnOffset = columnOffset ?? 0;
                ColumnHeaderRowHeight = columnHeaderRowHeight;
                ColumnHeaderStyle = columnHeaderStyle ?? new SLStyle();
                ShowAlternatingRows = showAlternatingRows ?? false;
                OddRowStyle = oddRowStyle ?? new SLStyle();
                EvenRowStyle = evenRowStyle ?? new SLStyle();
                UserDefinedColumns = userDefinedColumns ?? new List<Column>();
            }
            catch (Exception ex)
            {
                //WebLogger.LogException(
                //    new Exception(
                //        "SpreadsheetLightWrapper.Export.Models.ChildSetting.Contructor:Overload 4 -> " +
                //        ex.Message, ex),
                //    new Dictionary<string, string> {{"ChildSetting", "Constructor:Overload 4"}});
            }
        }

        #endregion Constructors

        #region User-Defined Column Functions

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     * Deprecated *
        ///     Alternative to creating a list of Column classes but with less functionality.
        ///     Overrides the default bound column names with custom ones and determines the formatting,
        ///     visibility and order.
        ///     <para />
        ///     "boundColumnName" string: Bound column name
        ///     <para />
        ///     "userDefinedColumnName" string: Custom column name
        ///     <para />
        ///     "numberFormat" string: Column format
        ///     <para />
        ///     "horizontalAlignment" string: Column horizontalAlignment
        ///     <para />
        ///     "showField" bool: Show/Hide column
        ///     <para />
        ///     "fieldOrder" int: * Optional - Set the order of the field
        ///     * Can be left out if showField is false
        ///     <para />
        ///     "userDefinedNumberFormat" string: * Optional - User-defined excel format
        ///     * Must be populated when "numberFormat" is set to "User-Defined"
        ///     <para />
        ///     ** Note: If none are added then the column names from the dataset table will be used and
        ///     there will no formatting.
        /// </summary>
        /// <param name="columnsDictionary">Dictionary(string, string): Dictionary of columns</param>
        /// -----------------------------------------------------------------------------------------------
        public void SetUserDefinedColumnNames(List<Column> columnsDictionary)
        {
            UserDefinedColumns = columnsDictionary;
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Overload 1: Overrides the default bound column names with custom ones and determines the formatting,
        ///     visibility and order.  In this context a Dictionary object is input where the only values
        ///     set are "boundColumnName" and "userDefinedColumnName"
        ///     <para />
        ///     "boundColumnName" string: Bound column name
        ///     <para />
        ///     "userDefinedColumnName" string: Custom column name
        ///     <para />
        ///     "numberFormat"  * Cannot be set
        ///     <para />
        ///     "horizontalAlignment"     * Cannot be set
        ///     <para />
        ///     "showField"          * Cannot be set	However, If the column is listed in dictionary input
        ///     then will be set to true, otherwise it will be set to false
        ///     <para />
        ///     "fieldOrder"    * Cannot be set	However, It is the order with which the fields are
        ///     listed in the Dictionary object.
        ///     <para />
        ///     "userDefinedNumberFormat" * Cannot be set
        /// </summary>
        /// <param name="columnsDictionary">Dictionary(string, string): Dictionary of columns</param>
        /// -----------------------------------------------------------------------------------------------
        public void SetUserDefinedColumnNames(Dictionary<string, string> columnsDictionary)
        {
            try
            {
                var i = 1;
                foreach (var row in columnsDictionary)
                {
                    if (UserDefinedColumns.Any(x => x.BoundColumnName == row.Key))
                    {
                        UserDefinedColumns.RemoveAt(UserDefinedColumns.FindIndex(x => x.BoundColumnName == row.Key));
                        UserDefinedColumns.Add(new Column
                        {
                            BoundColumnName = row.Key,
                            UserDefinedColumnName = row.Value,
                            ShowField = true,
                            FieldOrder = i
                        });
                    }
                    else
                        UserDefinedColumns.Add(new Column
                        {
                            BoundColumnName = row.Key,
                            UserDefinedColumnName = row.Value,
                            ShowField = true,
                            FieldOrder = i
                        });
                    i++;
                }
            }
            catch (Exception ex)
            {
                //WebLogger.LogException(
                //    new Exception(
                //        "SpreadsheetLightWrapper.Export.Models.ChildSetting.SetUserDefinedColumnNames -> " +
                //        ex.Message, ex),
                //    new Dictionary<string, string> {{"ChildSetting", "SetUserDefinedColumnNames"}});
            }
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Overload 2: Overrides the default bound column names with custom ones and determines the formatting,
        ///     visibility and order.
        ///     ** Note: If none are added then the column names from the dataset table will be used and
        ///     there will no formatting.
        ///     <para />
        ///     "boundColumnName" string: Bound column name
        ///     <para />
        ///     "userDefinedColumnName" string: Custom column name
        ///     <para />
        ///     "numberFormat" NumberFormats: Column format
        ///     <para />
        ///     "horizontalAlignment" HorizontalAlignmentValues: Column horizontalAlignment
        ///     <para />
        ///     "showField" bool: Show/Hide column
        ///     <para />
        ///     "fieldOrder" int: Set the order of the field
        ///     <para />
        ///     "userDefinedNumberFormat" string: User-defined excel format
        ///     * Must be populated when "numberFormat" is set to "User-Defined"
        /// </summary>
        /// <param name="boundColumnName">string: Bound column name</param>
        /// <param name="userDefinedColumnName">string: User-Defined column name</param>
        /// <param name="numberFormat">string: Column number format</param>
        /// <param name="horizontalAlignment">string: Column horizontalAlignment</param>
        /// <param name="showField">bool: Show/Hide column</param>
        /// <param name="fieldOrder">int: Set the order of the field</param>
        /// <param name="userDefinedNumberFormat">string</param>
        /// -----------------------------------------------------------------------------------------------
        public void SetUserDefinedColumnNames(
            string boundColumnName,
            string userDefinedColumnName,
            NumberFormats numberFormat,
            HorizontalAlignmentValues horizontalAlignment,
            bool showField,
            int? fieldOrder,
            string userDefinedNumberFormat = null)
        {
            try
            {
                if (UserDefinedColumns.Any(x => x.BoundColumnName == boundColumnName))
                {
                    UserDefinedColumns.RemoveAt(UserDefinedColumns.FindIndex(x => x.BoundColumnName == boundColumnName));
                    UserDefinedColumns.Add(new Column
                    {
                        BoundColumnName = boundColumnName,
                        UserDefinedColumnName = userDefinedColumnName,
                        NumberFormat = numberFormat,
                        HorizontalAlignment = horizontalAlignment,
                        ShowField = showField,
                        FieldOrder = fieldOrder,
                        UserDefinedNumberFormat = userDefinedNumberFormat
                    });
                }
                else
                    UserDefinedColumns.Add(new Column
                    {
                        BoundColumnName = boundColumnName,
                        UserDefinedColumnName = userDefinedColumnName,
                        NumberFormat = numberFormat,
                        HorizontalAlignment = horizontalAlignment,
                        ShowField = showField,
                        FieldOrder = fieldOrder,
                        UserDefinedNumberFormat = userDefinedNumberFormat
                    });
            }
            catch (Exception ex)
            {

            }
        }

        #endregion User-Defined Column Functions
    }
}