using System;
using System.Collections.Generic;
using System.Reflection;
using DocumentFormat.OpenXml.Spreadsheet;
using log4net;
using SpreadsheetLightWrapper.Core.style;
using SpreadsheetLightWrapper.Export;
using SpreadsheetLightWrapper.Export.Enums;
using SpreadsheetLightWrapper.Export.Models;
using SpreadsheetLightWrapper.Web.Mocks;
using Color = System.Drawing.Color;
using Column = SpreadsheetLightWrapper.Export.Models.Column;
using Page = System.Web.UI.Page;

namespace SpreadsheetLightWrapper.Web
{
    /// ===========================================================================================
    /// <summary>
    ///     Default Webpage to Demo the SpreadsheetLightWrapper Utility
    /// </summary>
    /// ===========================================================================================
    public partial class Default : Page
    {
        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Internal Members
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        private static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        private readonly MockDataCreator _mocks;

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Base Constructor
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        public Default()
        {
            _mocks = new MockDataCreator();
            /* Diagnostic */
            //Log.Info("Entering Default Constructor.");
            //ErrorHandlingTester(0);
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Base Page Load
        /// </summary>
        /// <param name="sender">object</param>
        /// <param name="e">EventArgs</param>
        /// -----------------------------------------------------------------------------------------------
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
            }
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Exports a basic grouped dataset that with four levels of drill-down; Directors, Managers
        ///     Team Leads & Associates
        ///     There are no User-defined columns and the DefaultSettings are used from the
        ///     DefaultExcelExportStyles static class in the App_Code folder.
        ///     In this case the data is stringified with no excel formatting.
        /// </summary>
        /// <param name="sender">object</param>
        /// <param name="e">EventArgs</param>
        /// -----------------------------------------------------------------------------------------------
        protected void btnBasicRelatedGroupedDataSet_Click(object sender, EventArgs e)
        {
            try
            {
                var dataSet = _mocks.CreateRelatedGroupedDataSet();
                Response.Clear();
                Response.ContentType = "application/excel";
                Response.AddHeader("Content-disposition", "filename=ExcelExport.xlsx");
                Exporter.OutputWorkbook(Response.OutputStream, dataSet, null,
                    DefaultExcelExportStyles.SetupDefaultStyles());
                Response.OutputStream.Flush();
                Response.OutputStream.Close();
                Response.Flush();
                Response.Close();
            }
            catch (Exception ex)
            {
                Log.Error("SpreadsheetLightWrapper.Web.Default.btnBasicRelatedGroupedDataSet_Click -> " + ex.Message +
                          ": " + ex);
            }
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Exports a basic grouped dataset that with four levels of drill-down; Directors, Managers
        ///     Team Leads & Associates
        ///     With this example there are User-Defined Columns for every dataset.
        /// </summary>
        /// <param name="sender">object</param>
        /// <param name="e">EventArgs</param>
        /// -----------------------------------------------------------------------------------------------
        protected void btnStyledRelatedGroupedDataSet_Click(object sender, EventArgs e)
        {
            try
            {
                var dataSet = _mocks.CreateRelatedGroupedDataSet();
                Response.Clear();
                Response.ContentType = "application/excel";
                Response.AddHeader("Content-disposition", "filename=ExcelExport.xlsx");
                Exporter.OutputWorkbook(
                    Response.OutputStream,
                    dataSet,
                    new[] {"Custom Grouped DS"},
                    CustomExcelExportStyles.SetupCustomStyles());
                Response.OutputStream.Flush();
                Response.OutputStream.Close();
                Response.Flush();
                Response.Close();
            }
            catch (Exception ex)
            {
                Log.Error("SpreadsheetLightWrapper.Web.Default.btnStyledRelatedGroupedDataSet_Click -> " + ex.Message +
                          ": " + ex);
            }
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Exports a basic grouped dataset that with four levels of drill-down; Directors, Managers
        ///     Team Leads & Associates
        ///     With this example there are User-Defined Columns for every dataset.
        ///     Saved to a file on C drive.
        /// </summary>
        /// <param name="sender">object</param>
        /// <param name="e">EventArgs</param>
        /// -----------------------------------------------------------------------------------------------
        protected void btnStyledRelatedGroupedDataSetToFile_Click(object sender, EventArgs e)
        {
            try
            {
                // Target path & filename
                var savePath = @"C:\\SpreadsheetLightWorkbook.xls";
                var dataSet = _mocks.CreateRelatedGroupedDataSet();
                Exporter.OutputWorkbook(
                    null,
                    dataSet,
                    new[] {"Custom Grouped DS"},
                    CustomExcelExportStyles.SetupCustomStyles(),
                    true,
                    savePath);
            }
            catch (Exception ex)
            {
                Log.Error("SpreadsheetLightWrapper.Web.Default.btnStyledRelatedGroupedDataSetToFile_Click -> " +
                          ex.Message + ": " + ex);
            }
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     This example has User-defined columns and Customized Settings. The User-Defined columns
        ///     set the column order, custom column names, visibility and data formatting.
        ///     It has four tables Directors, Managers, TeamLeads & Associates that are unrelated and
        ///     each gets its own sheet in the Workbook
        /// </summary>
        /// <param name="sender">object</param>
        /// <param name="e">EventArgs</param>
        /// -----------------------------------------------------------------------------------------------
        protected void btnUngroupedNoParentChildData_Click(object sender, EventArgs e)
        {
            try
            {
                var dataSet = _mocks.CreateUnrelatedUngroupedDataSet();
                Response.Clear();
                Response.ContentType = "application/excel";
                Response.AddHeader("Content-disposition", "filename=ExcelExport.xlsx");
                Exporter.OutputWorkbook(
                    Response.OutputStream,
                    dataSet,
                    null,
                    CustomExcelExportStyles.SetupCustomStyles());
                Response.OutputStream.Flush();
                Response.OutputStream.Close();
                Response.Flush();
                Response.Close();
            }
            catch (Exception ex)
            {
                Log.Error("SpreadsheetLightWrapper.Web.Default.btnUngroupedNoParentChildData_Click -> " + ex.Message +
                          ": " + ex);
            }
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     This example has four tables with two that are related, TeamLeads & Associates.
        ///     The unrelated tables get their own sheets, the two related go on one sheet.
        /// </summary>
        /// <param name="sender">object</param>
        /// <param name="e">EventArgs</param>
        /// -----------------------------------------------------------------------------------------------
        protected void btnPartiallyRelatedGroupedData_Click(object sender, EventArgs e)
        {
            try
            {
                var dataSet = _mocks.CreatePartiallyRelatedGroupedDataSet();
                Response.Clear();
                Response.ContentType = "application/excel";
                Response.AddHeader("Content-disposition", "filename=ExcelExport.xlsx");
                Exporter.OutputWorkbook(
                    Response.OutputStream,
                    dataSet,
                    null,
                    CustomExcelExportStyles.SetupCustomStyles());
                Response.OutputStream.Flush();
                Response.OutputStream.Close();
                Response.Flush();
                Response.Close();
            }
            catch (Exception ex)
            {
                Log.Error("SpreadsheetLightWrapper.Web.Default.btnPartiallyRelatedGroupedData_Click -> " + ex.Message +
                          ": " + ex);
            }
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     This example has four tables with two that are related, Managers & TeamLeads.
        ///     The unrelated tables get their own sheets, the two related go on one sheet.
        /// </summary>
        /// <param name="sender">object</param>
        /// <param name="e">EventArgs</param>
        /// -----------------------------------------------------------------------------------------------
        protected void btnPartiallyRelatedGroupedDataVer2_Click(object sender, EventArgs e)
        {
            try
            {
                var dataSet = _mocks.CreatePartiallyRelatedGroupedDataSetVer2();
                Response.Clear();
                Response.ContentType = "application/excel";
                Response.AddHeader("Content-disposition", "filename=ExcelExport.xlsx");
                Exporter.OutputWorkbook(
                    Response.OutputStream,
                    dataSet,
                    new[] {"Directors", "Managers-TeamLeads", "Associates"},
                    CustomExcelExportStyles.SetupCustomStyles());
                Response.OutputStream.Flush();
                Response.OutputStream.Close();
                Response.Flush();
                Response.Close();
            }
            catch (Exception ex)
            {
                Log.Error("SpreadsheetLightWrapper.Web.Default.btnPartiallyRelatedGroupedDataVer2_Click -> " +
                          ex.Message + ": " + ex);
            }
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     This example has four tables with three that are related, Directors, Managers & TeamLeads.
        ///     The unrelated table Associates will get its own sheet, the three related will be grouped
        ///     on one sheet.
        /// </summary>
        /// <param name="sender">object</param>
        /// <param name="e">EventArgs</param>
        /// -----------------------------------------------------------------------------------------------
        protected void btnPartiallyRelatedGroupedDataVer3_Click(object sender, EventArgs e)
        {
            try
            {
                var dataSet = _mocks.CreatePartiallyRelatedGroupedDataSetVer3();
                Response.Clear();
                Response.ContentType = "application/excel";
                Response.AddHeader("Content-disposition", "filename=ExcelExport.xlsx");
                Exporter.OutputWorkbook(
                    Response.OutputStream,
                    dataSet,
                    new[] {"Dir-Manag-TLs", "Associates"},
                    CustomExcelExportStyles.SetupCustomStyles());
                Response.OutputStream.Flush();
                Response.OutputStream.Close();
                Response.Flush();
                Response.Close();
            }
            catch (Exception ex)
            {
                Log.Error("SpreadsheetLightWrapper.Web.Default.btnPartiallyRelatedGroupedDataVer3_Click -> " +
                          ex.Message + ": " + ex);
            }
        }

        #region Diagnostics

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Automatically generate an error for testing log4net
        /// </summary>
        /// <param name="divisor"></param>
        /// -----------------------------------------------------------------------------------------------
        private void ErrorHandlingTester(int divisor)
        {
            try
            {
                var result = 2 / divisor;
            }
            catch (Exception ex)
            {
                Log.Error("SpreadsheetLightWrapper.Web.Default.ErrorHandlingTester -> " + ex.Message + ": " + ex);
            }
        }

        #endregion Diagnostics
    }

    /// ===========================================================================================
    /// <summary>
    ///     User-Defined Stylings class for Examples
    /// </summary>
    /// ===========================================================================================
    public static class CustomExcelExportStyles
    {
        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Internal Members
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        private static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     This setting creator has user-defined styles and columns for the four data-tables in
        ///     mock data; Directors, Managers, Team Leads & Associates.
        ///     Displays a variety of ways to access the Export library with Constructor and
        ///     Property Dependency Injection.
        ///     ** Note:  If you’re going to use a lot of the optional setting features then is
        ///     it recommended that you use Property injection, while Constructor injection is also
        ///     available it is limited to the most common cases of parameter input.
        /// </summary>
        /// <returns>Settings: Custom Styling</returns>
        /// -----------------------------------------------------------------------------------------------
        public static Settings SetupCustomStyles()
        {
            try
            {
                var childList = new List<ChildSetting>();
                /* -------------------------------------------------------------
                 * In this example the various ways of construction and setup
                 * are explored, use the ones that best suit your requirements.
                 * Setup the column header base style for the child datasets
                 * Documentation for SLStyle class in the SpreadsheetLight
                 * documentation & examples.
                 * For every Enum typing a period after the name will bring up
                 * a drop-down of items to select.
                 * -----------------------------------------------------------*/
                var baseColumnHeaderStyle = new SLStyle();
                baseColumnHeaderStyle.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
                baseColumnHeaderStyle.SetVerticalAlignment(VerticalAlignmentValues.Center);
                baseColumnHeaderStyle.Fill.SetPattern(PatternValues.Solid, Color.DimGray, Color.White);
                baseColumnHeaderStyle.SetBottomBorder(BorderStyleValues.Medium, Color.Black);
                baseColumnHeaderStyle.SetTopBorder(BorderStyleValues.Medium, Color.Black);
                baseColumnHeaderStyle.SetVerticalBorder(BorderStyleValues.Medium, Color.Black);
                baseColumnHeaderStyle.Border.SetRightBorder(BorderStyleValues.Medium, Color.Black);
                baseColumnHeaderStyle.Border.SetLeftBorder(BorderStyleValues.Medium, Color.Black);
                baseColumnHeaderStyle.SetFont("Britannic Bold", 12);
                baseColumnHeaderStyle.SetFontColor(Color.White);
                baseColumnHeaderStyle.SetFontBold(true);

                /* -------------------------------------------------------------
                 * Setup the odd row style for the child datasets
                 * -----------------------------------------------------------*/
                var oddRowStyle = new SLStyle();
                oddRowStyle.SetHorizontalAlignment(HorizontalAlignmentValues.Left);
                oddRowStyle.SetVerticalAlignment(VerticalAlignmentValues.Center);
                oddRowStyle.Fill.SetPattern(PatternValues.Solid, Color.White, Color.Black);
                oddRowStyle.SetFont("Helvetica", 10);
                oddRowStyle.SetFontColor(Color.Black);

                /* -------------------------------------------------------------
                 * Setup the even row style derived from the odd,
                 * change only what is necessary.
                 * -----------------------------------------------------------*/
                var evenRowStyle = oddRowStyle.Clone();
                evenRowStyle.Fill.SetPattern(PatternValues.Solid, Color.WhiteSmoke, Color.Black);

                /*  ------------------------------------------------------------
                 *  Create the user-defined columns with property dependency
                 *  injection for the base dataset.
                 *  With this method hover the cursor over the property and
                 *  intellisense will show the comments for it.
                 *  ----------------------------------------------------------*/
                var columns = new List<Column>
                {
                    // Since this id column is not set to visible, you can just leave it out and it will be ignored
                    new Column
                    {
                        BoundColumnName = "DID",
                        UserDefinedColumnName = "ID",
                        NumberFormat = NumberFormats.General,
                        HorizontalAlignment = HorizontalAlignmentValues.Center,
                        ShowField = false,
                        FieldOrder = 0
                    },
                    new Column
                    {
                        BoundColumnName = "Name",
                        UserDefinedColumnName = "Director",
                        NumberFormat = NumberFormats.General,
                        HorizontalAlignment = HorizontalAlignmentValues.Left,
                        ShowField = true,
                        FieldOrder = 1
                    },
                    new Column
                    {
                        BoundColumnName = "Age",
                        UserDefinedColumnName = "Chronology",
                        NumberFormat = NumberFormats.Decimal0,
                        HorizontalAlignment = HorizontalAlignmentValues.Center,
                        ShowField = true,
                        FieldOrder = 2
                    },
                    new Column
                    {
                        BoundColumnName = "Income",
                        UserDefinedColumnName = "Compensation",
                        NumberFormat = NumberFormats.UserDefined,
                        HorizontalAlignment = HorizontalAlignmentValues.Right,
                        ShowField = true,
                        FieldOrder = 3,
                        UserDefinedNumberFormat = "_($* #,##0.00_);[Red]_($* (#,##0.00);_($* \"-\"??_);_(@_)"
                    },
                    new Column
                    {
                        BoundColumnName = "Member",
                        UserDefinedColumnName = "Member ?",
                        NumberFormat = NumberFormats.General,
                        HorizontalAlignment = HorizontalAlignmentValues.Center,
                        ShowField = true,
                        FieldOrder = 4
                    },
                    new Column
                    {
                        BoundColumnName = "Registered",
                        UserDefinedColumnName = "Date Registered",
                        NumberFormat = NumberFormats.DateShort5,
                        HorizontalAlignment = HorizontalAlignmentValues.Center,
                        ShowField = true,
                        FieldOrder = 5
                    }
                };

                /* -------------------------------------------------------------
                 * Define and style base child settings.
                 * This Child will always be present, it represents the
                 * primary dataset for every export and is not really a child.
                 * Using Property Injection Technique
                 * -----------------------------------------------------------*/
                childList.Add(new ChildSetting
                {
                    // Optional name
                    SheetName = "Directors",
                    // Set column visibility
                    ShowColumnHeader = true,
                    // Make the base column header row a little larger
                    // so it will stand out.  Value is in pixels
                    ColumnHeaderRowHeight = 25,
                    // Setup the style for Column Headers
                    ColumnHeaderStyle = baseColumnHeaderStyle,
                    // Row and Alternating Row Styles
                    // If set to false then the odd row style will be overall row style
                    ShowAlternatingRows = false,
                    // Setup the style for all rows
                    OddRowStyle = oddRowStyle,
                    EvenRowStyle = null,
                    // Add the user-defined columns
                    UserDefinedColumns = columns
                });

                /*  ------------------------------------------------------------
                 *  The first child column headers stylings will be derived
                 *  from the base, change only what needs to be changed.
                 *  ----------------------------------------------------------*/
                var firstColumnHeaderStyle = baseColumnHeaderStyle.Clone();
                firstColumnHeaderStyle.Fill.SetPattern(PatternValues.Solid, Color.DarkGray, Color.Black);
                firstColumnHeaderStyle.SetBottomBorder(BorderStyleValues.Thin, Color.DarkSlateGray);
                firstColumnHeaderStyle.SetTopBorder(BorderStyleValues.Thin, Color.DarkSlateGray);
                firstColumnHeaderStyle.SetVerticalBorder(BorderStyleValues.Thin, Color.DarkSlateGray);
                firstColumnHeaderStyle.Border.SetRightBorder(BorderStyleValues.Thin, Color.DarkSlateGray);
                firstColumnHeaderStyle.Border.SetLeftBorder(BorderStyleValues.Thin, Color.DarkSlateGray);
                firstColumnHeaderStyle.SetFont("Helvetica", 10);
                firstColumnHeaderStyle.SetFontColor(Color.Black);

                /*  ------------------------------------------------------------
                 *  Create the user-defined columns with constructor dependency
                 *  injection for the base dataset.
                 *  Hover the cursor over the property and intellisense will
                 *  show the comments for it.
                 *  ----------------------------------------------------------*/
                columns = new List<Column>
                {
                    new Column
                    (
                        "Name",
                        "Managers",
                        NumberFormats.General,
                        HorizontalAlignmentValues.Left,
                        true,
                        1
                    ),
                    new Column
                    (
                        "Age",
                        "Age",
                        NumberFormats.UserDefined,
                        HorizontalAlignmentValues.Center,
                        true,
                        2,
                        "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)"
                    ),
                    new Column
                    (
                        "Income",
                        "Compensation",
                        NumberFormats.Currency0Black,
                        HorizontalAlignmentValues.Right,
                        true,
                        3
                    ),
                    new Column
                    (
                        "Registered",
                        "Date Registered",
                        NumberFormats.DateShort1,
                        HorizontalAlignmentValues.Center,
                        true,
                        5
                    )
                };
                /* -------------------------------------------------------------
                 * Define and add the first child
                 * Using Constructor dependency injection
                 * -----------------------------------------------------------*/
                childList.Add(new ChildSetting
                (
                    "Managers", // SheetName
                    true, // Show Column Headers
                    1, // Column Offset to the Right
                    null, // Column Header Row Height
                    firstColumnHeaderStyle, // Column Header Style
                    false, // Show Alternating Rows, false will default to Odd
                    oddRowStyle, // Odd Row Style
                    null, // Even Row Style
                    columns // User-Defined Column (UDCs)
                ));

                /* -------------------------------------------------------------
                 * The second child column headers stylings will be derived
                 * from the first, change only what needs to be changed.
                 * -----------------------------------------------------------*/
                var secondColumnHeaderStyle = firstColumnHeaderStyle.Clone();
                secondColumnHeaderStyle.Fill.SetPattern(PatternValues.Solid, Color.CadetBlue, Color.White);
                secondColumnHeaderStyle.SetFontColor(Color.White);

                /* -------------------------------------------------------------
                 * Define and add the second child
                 * Using Constructor dependency injection
                 * -----------------------------------------------------------*/
                childList.Add(new ChildSetting(
                    "Team Leads", // SheetName
                    true, // Show Column Headers
                    2, // Column Offset to the Right
                    null, // Column Header Row Height
                    secondColumnHeaderStyle, // Column Header Style
                    false, // Show Alternating Rows, false will default to Odd
                    oddRowStyle, // Odd Row Style
                    null, // Even Row Style
                    new List<Column> // User-Defined Column (UDCs)
                    {
                        new Column("TLID", "Team Lead ID", NumberFormats.General, HorizontalAlignmentValues.Left, true,
                            6),
                        new Column("Registered", "Registration Date", NumberFormats.UserDefined,
                            HorizontalAlignmentValues.Center, true, 2, "d-mmm-yy"),
                        new Column("Name", "Team Leads", NumberFormats.General, HorizontalAlignmentValues.Left, true, 0),
                        new Column("Age", "How Old?", NumberFormats.Decimal0, HorizontalAlignmentValues.Center, true, 1),
                        new Column("Member", "Member?", NumberFormats.General, HorizontalAlignmentValues.Center, true, 3),
                        new Column("Income", "Income", NumberFormats.Accounting2Red, HorizontalAlignmentValues.Right,
                            true, 4),
                        new Column("MID", "Foreign Key", NumberFormats.General, HorizontalAlignmentValues.Right, false)
                    }
                ));

                /* -------------------------------------------------------------
                 * The third child column headers stylings will be derived
                 * from the first, change only what needs to be changed.
                 * -----------------------------------------------------------*/
                var thirdColumnHeaderStyle = firstColumnHeaderStyle.Clone();
                thirdColumnHeaderStyle.Fill.SetPattern(PatternValues.Solid, Color.Aqua, Color.Black);
                thirdColumnHeaderStyle.SetFont("Blackadder ITC", 11);
                thirdColumnHeaderStyle.SetFontColor(Color.Black);

                /* -------------------------------------------------------------
                 * Define and add the third child
                 * Constructor Injection on all
                 * -----------------------------------------------------------*/
                childList.Add(new ChildSetting("Associates", true, 3, 30, thirdColumnHeaderStyle, true, oddRowStyle,
                    evenRowStyle,
                    new List<Column>
                    {
                        new Column("Name", "Associate", NumberFormats.General, HorizontalAlignmentValues.Left, true, 0),
                        new Column("Registered", "Date", NumberFormats.TimeStamp124, HorizontalAlignmentValues.Left,
                            true, 3),
                        new Column("Member", "Member?", NumberFormats.General, HorizontalAlignmentValues.Center, true, 2),
                        new Column("Name", "Associate", NumberFormats.General, HorizontalAlignmentValues.Left, true, 0)
                    }
                ));

                /* -------------------------------------------------------------
                 * Setup and return the primary container for the child datasets
                 * Using either Parameter or Constructor Injection.
                 * -----------------------------------------------------------*/
                var settings = new Settings
                {
                    Name = "Organization",
                    ChildSettings = childList
                };
                return settings;
            }
            catch (Exception ex)
            {
                Log.Error("SpreadsheetLightWrapper.Web.CustomExcelExportStyles -> " + ex.Message + ": " + ex);
            }
            return null;
        }
    }
}