using System;
using System.Data;
using System.IO;
using System.Linq;
using SpreadsheetLightWrapper.Core;
using SpreadsheetLightWrapper.Core.style;
using SpreadsheetLightWrapper.Export.Enums;
using SpreadsheetLightWrapper.Export.Models;

namespace SpreadsheetLightWrapper.Export
{
    /// ===========================================================================================
    /// <summary>
    ///     This a utility that enables DataSets with one or more tables be exported to Excel
    ///     Multiple tables must be linked with a primary key to foreign key relationship
    ///     which will be presented in an Excel Outline Group that can be expanded or collapsed.
    /// </summary>
    /// ===========================================================================================
    public class Exporter
    {
        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Internal Members
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        private static Settings _settings;

        private static int _tableCounter;
        private static int _sheetCounter;
        private readonly SLDocument _document;

        #region Properties

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Properties
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        public Settings Settings
        {
            get { return _settings; }
            set { _settings = value; }
        }

        #endregion Properties

        #region Constructors

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Overload 1: Base Constructor
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        public Exporter()
        {
            try
            {
                _settings = DefaultStyles.SetupDefaultStyles();
                _document = new SLDocument();
            }
            catch (Exception ex)
            {
                //WebLogger.LogException(
                //    new Exception(
                //        "Ups.Toolkit.SpreadsheetLight.Exporter.Constructor: Overload 1 -> " + ex.Message, ex),
                //    new Dictionary<string, string> {{"Exporter", "Constructor: Overload 1"}});
            }
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Overload 2: Constructor
        /// </summary>
        /// <param name="settings">Settings</param>
        /// -----------------------------------------------------------------------------------------------
        private Exporter(Settings settings)
        {
            try
            {
                _document = new SLDocument();

                /*  ------------------------------------------------------------
                 *  Initialize the default styles
                 *  If the user doesn't input any then get the Default styles
                 *  from the Ups.Toolkit.SpreadsheetLight library
                 *  ----------------------------------------------------------*/
                _settings = settings ?? DefaultStyles.SetupDefaultStyles();
            }
            catch (Exception ex)
            {
                //WebLogger.LogException(
                //    new Exception(
                //        "Ups.Toolkit.SpreadsheetLight.Exporter.Constructor: Overload 2 -> " + ex.Message, ex),
                //    new Dictionary<string, string> {{"Exporter", "Constructor: Overload 2"}});
            }
        }

        #endregion Constructors

        #region Static Implementations Section

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Overload 1: Base function that generates the workbook structure in a factory pattern;
        ///     setting styles and values.
        ///     Takes a DataTable
        /// </summary>
        /// <param name="outputStream">Stream: Response.OutputStream</param>
        /// <param name="dataTable">dataTable</param>
        /// <param name="sheetNames">string[]: Names of the individual Sheets in your output</param>
        /// <param name="settings">Settings: User-Defined Settings (Optional)</param>
        /// <param name="showColumnHeaders">bool: Turn all Column Headers On/Off (Optional)</param>
        /// <param name="fileNameAndPath">string: Filename & Path if output is directed to specific location</param>
        /// -----------------------------------------------------------------------------------------------
        public static void OutputWorkbook(
            Stream outputStream,
            DataTable dataTable,
            string[] sheetNames = null,
            Settings settings = null,
            bool showColumnHeaders = true,
            string fileNameAndPath = null)
        {
            try
            {
                // Setup the Helper Class
                var exporter = new Exporter(settings);
                // Add Data then add it to the new DataSet
                var dataSet = new DataSet();
                dataSet.Tables.Add(dataTable);
                var bk = exporter.GenerateWorkbook(dataSet, sheetNames, showColumnHeaders);
                // Saving Workbook to file path or not
                if (fileNameAndPath == null)
                    bk.SaveAs(outputStream);
                else
                    bk.SaveAs(fileNameAndPath);
            }
            catch (Exception ex)
            {
                //WebLogger.LogException(
                //    new Exception(
                //        "Ups.Toolkit.SpreadsheetLight.Exporter.OutputWorkbook:Overload 1 -> " + ex.Message, ex),
                //    new Dictionary<string, string> {{"Exporter", "OutputWorkbook:Overload 1"}});
            }
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Overload 2: Base function that generates the workbook structure in a factory pattern;
        ///     setting styles and values.
        ///     Takes a DataSet.
        /// </summary>
        /// <param name="outputStream">Stream</param>
        /// <param name="dataSet">DataSet</param>
        /// <param name="sheetNames">string[]: Names of the individual Sheets in your output</param>
        /// <param name="settings">Settings: User-Defined Settings (Optional)</param>
        /// <param name="showColumnHeaders">bool: Turn all Column Headers On/Off (Optional)</param>
        /// <param name="fileNameAndPath">string: Filename & Path if output is directed to specific location</param>
        /// -----------------------------------------------------------------------------------------------
        public static void OutputWorkbook(
            Stream outputStream,
            DataSet dataSet,
            string[] sheetNames = null,
            Settings settings = null,
            bool showColumnHeaders = true,
            string fileNameAndPath = null)
        {
            try
            {
                // Setup the Helper Class
                var exporter = new Exporter(settings);
                // Add Data
                var bk = exporter.GenerateWorkbook(dataSet, sheetNames, showColumnHeaders);
                // Saving Workbook to file path or not
                if (fileNameAndPath == null)
                    bk.SaveAs(outputStream);
                else
                    bk.SaveAs(fileNameAndPath);
            }
            catch (Exception ex)
            {
                //WebLogger.LogException(
                //    new Exception(
                //        "Ups.Toolkit.SpreadsheetLight.Exporter.OutputWorkbook:Overload 2 -> " + ex.Message, ex),
                //    new Dictionary<string, string> {{"Exporter", "OutputWorkbook:Overload 2 "}});
            }
        }

        #endregion Static Implementations Section

        #region Utilities

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Get Workbook, setup and insert worksheet, add data, add styles
        /// </summary>
        /// <param name="dataSet">DataSet</param>
        /// <param name="sheetNames">string[]: Names of the individual Sheets in your output</param>
        /// <param name="showColumnHeadersOverall">bool</param>
        /// <param name="rowIndex">int</param>
        /// <param name="columnIndex">int</param>
        /// <returns>SLDocument</returns>
        /// -----------------------------------------------------------------------------------------------
        private SLDocument GenerateWorkbook(
            DataSet dataSet,
            string[] sheetNames,
            bool showColumnHeadersOverall,
            int rowIndex = 1,
            int columnIndex = 1)
        {
            try
            {
                // Declarations
                var tableCounter = 0;
                var outlineLevel = 0;
                // Initializtions
                _tableCounter = tableCounter;
                _sheetCounter = 0;

                // Get first table from Dataset
                var dataTable = dataSet.Tables[tableCounter];

                // If not then get first child
                var baseChild = _settings.ChildSettings[tableCounter];
                // Get a name for this sheet
                var workingSheetName = GetSheetName(
                    _sheetCounter,
                    sheetNames,
                    baseChild.SheetName,
                    dataTable.TableName);

                // Set the sheet name and select it.
                _document.RenameWorksheet(SLDocument.DefaultFirstSheetName, workingSheetName);
                _document.SelectWorksheet(workingSheetName);

                // Set the row grouping +/- graphic to appear at the top of the grouping
                _document.slws.PageSettings.SheetProperties.SummaryBelow = false;
                // Set the column grouping +/- graphic to appear at the right of the grouping
                _document.slws.PageSettings.SheetProperties.SummaryRight = false;

                // Get rows and relation count
                var dataRowCollection = dataSet.Tables[tableCounter].Rows;

                // Get the bound columns
                var columns = dataTable.Columns;

                // Get the UDC sorted by column order
                var sortedUdc = baseChild.UserDefinedColumns.OrderBy(item => item.FieldOrder);

                // Are the column headers going to be shown?
                // First check "showColumnHeadersOverall" to see if they have all been turned off (set to false),
                // if not then get individual child setting.
                var showColumnHeaders = showColumnHeadersOverall && baseChild.ShowColumnHeader;

                // Initialize rowCounter counter to starting rowCounter index;
                var rowCounter = rowIndex;

                // If no records append "No Data" to sheet and return document
                // otherwise proceed with data export
                if (dataRowCollection.Count == 0)
                {
                    _document.SetCellValue(rowCounter, columnIndex + 1, "No Data");
                }
                else
                {
                    int childRelationCount;
                    /****** Adding Column Headers *******/
                    // Add column headers from either the Data-table or from the UDCs
                    SetupColumnHeaders(columnIndex, ref rowCounter, showColumnHeaders, outlineLevel, sortedUdc,
                        columns,
                        baseChild);

                    /****** Adding Parent Content Rows *******/
                    //Looping through Data Table content
                    var rowOdd = true; // First rowCounter is odd
                    foreach (DataRow parentRow in dataRowCollection)
                    {
                        //Alternate Rows Style
                        SetupRowsAndCells(rowCounter, columnIndex, rowOdd, sortedUdc, columns, parentRow,
                            tableCounter);
                        rowOdd = !rowOdd; // odd/even switch

                        // Increment the rowCounter count
                        ++rowCounter;

                        // Is there a Child Relation
                        childRelationCount = dataSet.Tables[tableCounter].ChildRelations.Count;
                        // If so, begin the recursive call for the children
                        if (childRelationCount != 0)
                        {
                            // Get the child relation name and call for its children, increment tableCounter & outlineLevel
                            // for the next family of children
                            var relationName = dataSet.Tables[tableCounter].ChildRelations[0].ToString();
                            GetChildren(dataSet, parentRow, workingSheetName, tableCounter + 1, outlineLevel + 1,
                                relationName, showColumnHeadersOverall, ref rowCounter, columnIndex);
                        }
                    }
                    _document.SelectWorksheet(workingSheetName);
                    _document.AutoFitColumn(1, 30);
                    _document.AutoFitRow(1, rowCounter);

                    // Is there there another table and does it have a Child Relation?
                    childRelationCount = dataSet.Tables[_tableCounter].ChildRelations.Count;
                    if (_tableCounter + 1 < dataSet.Tables.Count && childRelationCount == 0)
                        GetSubsequentSheets(dataSet, sheetNames, _tableCounter + 1, showColumnHeadersOverall);
                }
                return _document;
            }
            catch (Exception ex)
            {
                //WebLogger.LogException(
                //    new Exception(
                //        "Ups.Toolkit.SpreadsheetLight.Exporter.GenerateWorkbook -> " + ex.Message, ex),
                //    new Dictionary<string, string> {{"Exporter", "GenerateWorkbook"}});
            }
            return null;
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Recursively called function that adds another sheet  formatting and column headers
        /// </summary>
        /// <param name="dataSet">DataSet</param>
        /// <param name="sheetNames">string[]: Names of the individual Sheets in your output</param>
        /// <param name="tableCounter">int</param>
        /// <param name="showColumnHeadersOverall">bool</param>
        /// -----------------------------------------------------------------------------------------------
        private void GetSubsequentSheets(
            DataSet dataSet,
            string[] sheetNames,
            int tableCounter,
            bool showColumnHeadersOverall)
        {
            try
            {
                // Declarations
                var outlineLevel = 0;
                var rowIndex = 1;
                var columnIndex = 1;

                // Get first table from Dataset
                var dataTable = dataSet.Tables[tableCounter];
                // Increment the overall table counter
                _tableCounter = _tableCounter < tableCounter ? tableCounter : _tableCounter;

                // If not then get first child
                var baseChild = _settings.ChildSettings[tableCounter];

                // Increment the overall sheet counter to get the next user-defined name if it exists
                // Get a name for this sheet
                var workingSheetName = GetSheetName(
                    ++_sheetCounter,
                    sheetNames,
                    baseChild.SheetName,
                    dataTable.TableName);

                // Create a new sheet, name it & select it.
                _document.AddWorksheet(workingSheetName);
                _document.SelectWorksheet(workingSheetName);

                // Set the row grouping +/- graphic to appear at the top of the grouping
                _document.slws.PageSettings.SheetProperties.SummaryBelow = false;
                // Set the column grouping +/- graphic to appear at the right of the grouping
                _document.slws.PageSettings.SheetProperties.SummaryRight = false;

                // Get rows and relation count
                var dataRowCollection = dataSet.Tables[tableCounter].Rows;

                // Get the bound columns
                var columns = dataTable.Columns;

                // Get the UDC sorted by column order
                var sortedUdc = baseChild.UserDefinedColumns.OrderBy(item => item.FieldOrder);

                // Are the column headers going to be shown?
                // First check "showColumnHeadersOverall" to see if they have all been turned off (set to false),
                // if not then get individual child setting.
                var showColumnHeaders = showColumnHeadersOverall && baseChild.ShowColumnHeader;

                // Initialize rowCounter counter to starting rowCounter index;
                var rowCounter = rowIndex;

                // If no records append "No Data" to sheet and return document
                // otherwise proceed with data export
                if (dataRowCollection.Count == 0)
                {
                    _document.SetCellValue(rowCounter, columnIndex + 1, "No Data");
                }
                else
                {
                    /****** Adding Column Headers *******/
                    // Add column headers from either the Data-table or from the UDCs
                    SetupColumnHeaders(columnIndex, ref rowCounter, showColumnHeaders, outlineLevel, sortedUdc, columns,
                        baseChild);

                    // Is there a Child Relation
                    var childRelationCount = dataSet.Tables[tableCounter].ChildRelations.Count;

                    /****** Adding Parent Content Rows *******/
                    //Looping through Data Table content
                    var rowOdd = true; // First rowCounter is odd
                    foreach (DataRow parentRow in dataRowCollection)
                    {
                        //Alternate Rows Style
                        SetupRowsAndCells(rowCounter, columnIndex, rowOdd, sortedUdc, columns, parentRow, tableCounter);
                        rowOdd = !rowOdd; // odd/even switch

                        // Increment the rowCounter count
                        ++rowCounter;

                        // If so, begin the recursive call for the children
                        if (childRelationCount != 0)
                        {
                            // Get the child relation name and call for its children, increment tableCounter & outlineLevel
                            // for the next family of children
                            var relationName = dataSet.Tables[tableCounter].ChildRelations[0].ToString();
                            GetChildren(dataSet, parentRow, workingSheetName, tableCounter + 1, outlineLevel + 1,
                                relationName,
                                showColumnHeadersOverall, ref rowCounter, columnIndex);
                        }
                    }

                    _document.SelectWorksheet(workingSheetName);
                    _document.AutoFitColumn(1, 30);
                    _document.AutoFitRow(1, rowCounter);
                    // Is there there another table and does it have a Child Relation?
                    childRelationCount = dataSet.Tables[_tableCounter].ChildRelations.Count;
                    if (_tableCounter + 1 < dataSet.Tables.Count && childRelationCount == 0)
                        GetSubsequentSheets(dataSet, sheetNames, _tableCounter + 1, showColumnHeadersOverall);
                }
            }
            catch (Exception ex)
            {
                //WebLogger.LogException(
                //    new Exception(
                //        "Ups.Toolkit.SL.Exporter.GetChildren -> " + ex.Message, ex),
                //    new Dictionary<string, string> {{"Exporter", "GetChildren"}});
            }
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Recursively called function that adds the child rows, formatting and column headers
        /// </summary>
        /// <param name="dataSet">DataSet</param>
        /// <param name="parentRow">DataRow</param>
        /// <param name="sheetName">string</param>
        /// <param name="tableCounter">int</param>
        /// <param name="outlineLevel">int</param>
        /// <param name="relationName">string</param>
        /// <param name="showColumnHeadersOverall">bool</param>
        /// <param name="rowCounter">int</param>
        /// <param name="columnIndex">int</param>
        /// -----------------------------------------------------------------------------------------------
        private void GetChildren(
            DataSet dataSet,
            DataRow parentRow,
            string sheetName,
            int tableCounter,
            int outlineLevel,
            string relationName,
            bool showColumnHeadersOverall,
            ref int rowCounter,
            int columnIndex)
        {
            try
            {
                // Get next child table from dataset
                var dataTable = dataSet.Tables[tableCounter];

                // Select the current sheet.
                _document.SelectWorksheet(sheetName);

                // Get the bound columns from data table
                var columns = dataTable.Columns;

                // Get the Child Stylings
                var childSetting = _settings.ChildSettings[tableCounter];

                // Get number of column offsets
                // Add to the starting column index to adjust dataset to the left
                var offsetColumnIndex = _settings.ChildSettings[tableCounter].ColumnOffset != 0
                    ? _settings.ChildSettings[tableCounter].ColumnOffset + columnIndex
                    : columnIndex;

                // Get the UDC sorted by column order
                var sortedUdc = childSetting.UserDefinedColumns.OrderBy(item => item.FieldOrder);

                // Get the child rows
                var children = parentRow.GetChildRows(relationName);

                // If there's any children for this parent print them out
                if (children.Length != 0)
                {
                    // Are the column headers going to be shown?
                    // First check "showColumnHeadersOverall" to see if they have all been turned off (set to false),
                    // if not then get individual child setting.
                    var showColumnHeaders = showColumnHeadersOverall && childSetting.ShowColumnHeader;

                    /****** Adding Column Headers *******/
                    // Add column headers from either the Data-table or from the UDCs
                    SetupColumnHeaders(offsetColumnIndex, ref rowCounter, showColumnHeaders, outlineLevel, sortedUdc,
                        columns, childSetting);

                    // Increment the overall table counter
                    _tableCounter = _tableCounter < tableCounter ? tableCounter : _tableCounter;

                    /****** Adding Content Child Table Rows *******/
                    // Looping through Data Table content
                    var rowOdd = true; // First rowCounter is odd
                    foreach (var child in children)
                    {
                        var colCounter = offsetColumnIndex;

                        //Alternate Rows Style
                        SetupRowsAndCells(rowCounter, colCounter, rowOdd, sortedUdc, columns, child, tableCounter);
                        rowOdd = !rowOdd; // odd/even switch

                        // Setup the grouping and Outline level for this child
                        _document.AddGroupedRow(rowCounter, outlineLevel);

                        // Increment the rowCounter count
                        ++rowCounter;

                        // Is there a Child Relation
                        var childRelationCount = dataSet.Tables[tableCounter].ChildRelations.Count;

                        // If so, begin the recursive call for the children
                        if (childRelationCount != 0)
                        {
                            // Get the child relation name and call for its children, increment tableCounter & outlineLevel
                            // for the next family of children
                            relationName = dataSet.Tables[tableCounter].ChildRelations[0].ToString();
                            GetChildren(dataSet, child, sheetName, tableCounter + 1, outlineLevel + 1, relationName,
                                showColumnHeadersOverall, ref rowCounter, columnIndex);
                        }
                    }
                    _document.CollapseAllRows(rowCounter);
                }
            }
            catch (Exception ex)
            {
                //WebLogger.LogException(
                //    new Exception(
                //        "Ups.Toolkit.SL.Exporter.GetChildren -> " + ex.Message, ex),
                //    new Dictionary<string, string> {{"Exporter", "GetChildren"}});
            }
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Determine a sheet name:
        ///     First look at the input sheet name.
        ///     Then look at the sheet name from the ChildSettings
        ///     Then look for a Table name to use
        ///     Finally a non duplication default value
        ///     ** Note: These characters cannot be used in a file name:
        ///     greater-than or less-than signs, asterisks, question marks, double quotes,
        ///     vertical bars or pipes, colons, forward slashes or brackets.
        /// </summary>
        /// <param name="sheetCounter">int</param>
        /// <param name="sheetNames">string[]: Names of the individual Sheets in your output</param>
        /// <param name="childSheetName">string</param>
        /// <param name="tableSheetName">string</param>
        /// <returns>string: Working Sheet Name</returns>
        /// -----------------------------------------------------------------------------------------------
        private string GetSheetName(int sheetCounter, string[] sheetNames = null, string childSheetName = null,
            string tableSheetName = null)
        {
            // List of invalid characters for Sheet Name
            var badSheetNameChar = new[] {'<', '>', '*', '?', '"', '|', ':', '/', '[', ']'};
            string workingSheetName;
            // Any user-defined sheet names
            if (sheetNames != null)
            {
                // Get the user defined sheet name
                workingSheetName = sheetNames[sheetCounter];
                // Has the name got any invalid characters for a Sheet Name
                if (!BadCharacters(workingSheetName, badSheetNameChar))
                    return "Bad Character in Name"; // return feedback for the developer
            }
            else
            {
                // Otherwise get the default names
                if (!string.IsNullOrEmpty(childSheetName))
                    // Check for a child SheetName then capture it.
                    workingSheetName = childSheetName;
                else
                    workingSheetName = !string.IsNullOrEmpty(tableSheetName)
                        ? tableSheetName
                        : "output" + sheetCounter;
            }
            return workingSheetName;
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Checks for invalid characters in a string
        /// </summary>
        /// <param name="name">string: The input string to be tested</param>
        /// <param name="invalidCharacters">char[]: List of character that cannot be in an input string</param>
        /// <returns></returns>
        /// -----------------------------------------------------------------------------------------------
        private bool BadCharacters(string name, char[] invalidCharacters)
        {
            foreach (var item in invalidCharacters)
                if (name.Contains(item))
                   return false;
            return true;
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Add column headers from either the Data-table or from the UDCs and format them
        /// </summary>
        /// <param name="offsetColumnIndex">int</param>
        /// <param name="rowCounter">ref int</param>
        /// <param name="showColumnHeaders">bool</param>
        /// <param name="outlineLevel">int</param>
        /// <param name="sortedUdc">IOrderedEnumerable(Column)</param>
        /// <param name="columns">DataColumnCollection</param>
        /// <param name="childSetting">ChildSetting</param>
        /// -----------------------------------------------------------------------------------------------
        private void SetupColumnHeaders(
            int offsetColumnIndex,
            ref int rowCounter,
            bool showColumnHeaders,
            int outlineLevel,
            IOrderedEnumerable<Column> sortedUdc,
            DataColumnCollection columns,
            ChildSetting childSetting)
        {
            if (showColumnHeaders)
            {
                // Initialize column counter to starting column index;
                var colCounter = offsetColumnIndex;
                // If there aren't any then pass the bound column headers through
                if (!sortedUdc.Any())
                    foreach (DataColumn col in columns)
                    {
                        // Add the default column name from Data-table
                        _document.SetCellValue(rowCounter, colCounter, col.ColumnName);
                        _document.SetCellStyle(rowCounter, colCounter, childSetting.ColumnHeaderStyle);
                        ++colCounter;
                    }
                else
                    foreach (var sudc in sortedUdc)
                    foreach (DataColumn col in columns)
                        if (sudc.BoundColumnName == col.ColumnName)
                            if (sudc.ShowField)
                            {
                                // Add the UDC (User Defined Column) name
                                _document.SetCellValue(rowCounter, colCounter, sudc.UserDefinedColumnName);
                                _document.SetCellStyle(rowCounter, colCounter, childSetting.ColumnHeaderStyle);
                                ++colCounter;
                            }
                // Set grouping level for this child's headers
                _document.AddGroupedRow(rowCounter, outlineLevel);
                // Set column header height
                if (childSetting.ColumnHeaderRowHeight != null)
                    _document.SetRowHeight(rowCounter, (double) childSetting.ColumnHeaderRowHeight);
                // Since column headers were added then increment the rowCounter before it's returned by reference
                ++rowCounter;
            }
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     This function creates and adds cells to the rowCounter and formats them along with overall
        ///     styles for the rowCounter by calling the function "SetRowAndCellStyles".
        /// </summary>
        /// <param name="row">int</param>
        /// <param name="columnIndex">int</param>
        /// <param name="rowOdd">bool: odd or even rowCounter</param>
        /// <param name="sortedUdc">IOrderedEnumerable(Column): Sorted user defined columns</param>
        /// <param name="columns">DataColumnCollection</param>
        /// <param name="dr">DataRow: current data-row</param>
        /// <param name="tableCounter">int</param>
        /// -----------------------------------------------------------------------------------------------
        private void SetupRowsAndCells(
            int row,
            int columnIndex,
            bool rowOdd,
            IOrderedEnumerable<Column> sortedUdc,
            DataColumnCollection columns,
            DataRow dr,
            int tableCounter)
        {
            try
            {
                SLStyle style;
                // Are alternating rowCounter styles to be used?
                // If true set accordingly
                if (_settings.ChildSettings[tableCounter].ShowAlternatingRows)
                    style = rowOdd
                        ? _settings.ChildSettings[tableCounter].OddRowStyle
                        : _settings.ChildSettings[tableCounter].EvenRowStyle;
                else
                    style = _settings.ChildSettings[tableCounter].OddRowStyle;

                // If no User-Defined Columns, then the field values will be output as is
                if (!sortedUdc.Any())
                {
                    // Reset column to starting column index;
                    var colCounter = columnIndex;
                    // Add all the fields from the data-table
                    foreach (DataColumn col in columns)
                    {
                        // Since no UDC data formats stringify everything
                        _document.SetCellValue(row, colCounter, dr[col].ToString());
                        // Set the basic style from the settings classes
                        _document.SetCellStyle(row, colCounter, style);
                        // On to the next row
                        ++colCounter;
                    }
                }
                else
                {
                    // Reset column to starting column index;
                    var colCounter = columnIndex;
                    // Then loop through sorted UDC
                    foreach (var sudc in sortedUdc)
                    foreach (DataColumn col in columns)
                        if (sudc.BoundColumnName == col.ColumnName)
                            if (sudc.ShowField)
                            {
                                // Set data type so excel number formats will work
                                SetDataType(row, dr, col, colCounter);
                                // Get left, center or right alignment from UDCs
                                style.Alignment.Horizontal = sudc.HorizontalAlignment;
                                // Get the number formats from the UDCs
                                if (sudc.NumberFormat == NumberFormats.UserDefined)
                                    // If there is a user-defined format then add the extra field
                                    // with the custom excel format
                                    SetNumberFormat(sudc.NumberFormat, ref style, sudc.UserDefinedNumberFormat);
                                else
                                    // Otherwise use the predefined ones
                                    SetNumberFormat(sudc.NumberFormat, ref style);
                                // Set the basic style from the settings classes
                                _document.SetCellStyle(row, colCounter, style);
                                // On to the next row
                                ++colCounter;
                            }
                }
            }
            catch (Exception ex)
            {
                //WebLogger.LogException(
                //    new Exception(
                //        "Ups.Toolkit.SpreadsheetLight.Exporter.SetupRowsAndCells -> " + ex.Message, ex),
                //    new Dictionary<string, string> {{"Exporter", "SetupRowsAndCells"}});
            }
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Set the numeric format of the cell text value from a predetermined
        ///     list of formats enumerated in the NumberFormats Enum.
        ///     If format is "User-Defined" then add a valid Excel Number Format in "userDefinedExcelFormat"
        /// </summary>
        /// <param name="format">NumberFormats</param>
        /// <param name="style">ref SLStyle</param>
        /// <param name="userDefinedNumberFormat">string</param>
        /// -----------------------------------------------------------------------------------------------
        private void SetNumberFormat(NumberFormats format, ref SLStyle style, string userDefinedNumberFormat = null)
        {
            try
            {
                style.FormatCode = format == NumberFormats.UserDefined ? userDefinedNumberFormat : GetAttribute(format);
            }
            catch (Exception ex)
            {
                //WebLogger.LogException(
                //    new Exception(
                //        "Ups.Toolkit.SpreadsheetLight.Exporter.SetNumberFormat -> " + ex.Message, ex),
                //    new Dictionary<string, string> {{"Exporter", "SetNumberFormat"}});
            }
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Gets the Format String attribute from the Enum Value
        /// </summary>
        /// <param name="format">NumberFormats</param>
        /// <returns>string</returns>
        /// -----------------------------------------------------------------------------------------------
        private string GetAttribute(NumberFormats format)
        {
            try
            {
                var type = format.GetType();
                var fi = type.GetField(format.ToString());
                var formatString = fi.GetCustomAttributes(typeof(FormatString), false) as FormatString[];
                if (formatString != null) return formatString[0].Value;
            }
            catch (Exception ex)
            {
                //WebLogger.LogException(
                //    new Exception(
                //        "Ups.Toolkit.SpreadsheetLight.Exporter.GetAttribute -> " + ex.Message, ex),
                //    new Dictionary<string, string> {{"Exporter", "GetAttribute"}});
            }
            return null;
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Determines the Data Type and casts the value into "SetCellValue"
        /// </summary>
        /// <param name="row">int</param>
        /// <param name="dr">DataRow</param>
        /// <param name="col">DataColumn</param>
        /// <param name="colCounter">int</param>
        /// -----------------------------------------------------------------------------------------------
        private void SetDataType(int row, DataRow dr, DataColumn col, int colCounter)
        {
            try
            {
                switch (col.DataType.Name)
                {
                    case "Byte":
                        _document.SetCellValue(row, colCounter, (byte) dr[col]);
                        break;
                    case "Boolean":
                        _document.SetCellValue(row, colCounter, (bool) dr[col]);
                        break;
                    case "DateTime":
                        _document.SetCellValue(row, colCounter, (DateTime) dr[col]);
                        break;
                    case "Int16":
                        _document.SetCellValue(row, colCounter, (short) dr[col]);
                        break;
                    case "Int32":
                        _document.SetCellValue(row, colCounter, (int) dr[col]);
                        break;
                    case "Int64":
                        _document.SetCellValue(row, colCounter, (long) dr[col]);
                        break;
                    case "UInt16":
                        _document.SetCellValue(row, colCounter, (ushort) dr[col]);
                        break;
                    case "UInt32":
                        _document.SetCellValue(row, colCounter, (uint) dr[col]);
                        break;
                    case "UInt64":
                        _document.SetCellValue(row, colCounter, (ulong) dr[col]);
                        break;
                    case "Double":
                        _document.SetCellValue(row, colCounter, (double) dr[col]);
                        break;
                    case "Float":
                        _document.SetCellValue(row, colCounter, (float) dr[col]);
                        break;
                    case "Decimal":
                        _document.SetCellValue(row, colCounter, (decimal) dr[col]);
                        break;
                    default:
                        // All other data types; String, DBNull, etc.
                        _document.SetCellValue(row, colCounter, dr[col].ToString());
                        break;
                }
            }
            catch (Exception ex)
            {
                //WebLogger.LogException(
                //    new Exception(
                //        "Ups.Toolkit.SpreadsheetLight.Exporter.SetDataType -> " + ex.Message, ex),
                //    new Dictionary<string, string> {{"Exporter", "SetDataType"}});
            }
        }

        #endregion Utilities
    }
}