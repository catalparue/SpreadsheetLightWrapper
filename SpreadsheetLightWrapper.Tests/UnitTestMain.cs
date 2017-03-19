using System;
using System.Collections.Generic;
using System.Data;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SpreadsheetLightWrapper.Core.style;
using SpreadsheetLightWrapper.Export;
using SpreadsheetLightWrapper.Export.Enums;
using SpreadsheetLightWrapper.Export.Models;
using Color = System.Drawing.Color;
using Column = SpreadsheetLightWrapper.Export.Models.Column;

namespace SpreadsheetLightWrapper.Tests
{
    /// ===========================================================================================
    /// <summary>
    ///     Test class for SpreadsheetLightWrapper
    /// </summary>
    /// ===========================================================================================
    [TestClass]
    public class UnitTestMain
    {
        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Test the basic library
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        [TestMethod]
        public void ExportTest()
        {
            try
            {
                var savePath = @"C:\\SpreadsheetLightWorkbookSimpleSave.xls";
                var data = new CreateMockData();
                var dataSet = data.CreateDataSet();
                Exporter.OutputWorkbook(null, dataSet, new[] {"Custom Grouped DS"},
                    CustomExcelExportStyles.SetupCustomStyles(), true, savePath);
            }
            catch (Exception ex)
            {
                Assert.Fail("Exception Fail: " + ex.Message);
            }
        }
    }

    /// ===========================================================================================
    /// <summary>
    ///     User-Defined Styling class for Examples
    /// </summary>
    /// ===========================================================================================
    public static class CustomExcelExportStyles
    {
        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     This setting creator has user-defined styles and columns for the four data-tables in
        ///     mock data; Directors, Managers, Team Leads & Associates.
        ///     Displays a variety of ways to access the Export library with Constructor and
        ///     Property Dependency Injection.
        /// </summary>
        /// <returns>Settings: Custom Styling</returns>
        /// -----------------------------------------------------------------------------------------------
        public static Settings SetupCustomStyles()
        {
            try
            {
                var childList = new List<ChildSetting>();
                /* -------------------------------------------------------------
                 * Setup the column header base style for the child datasets
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
                        BoundColumnName = "SheetName",
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
                        NumberFormat = NumberFormats.Accounting2Red,
                        HorizontalAlignment = HorizontalAlignmentValues.Right,
                        ShowField = true,
                        FieldOrder = 3
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
                        "SheetName",
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
                        NumberFormats.Decimal0,
                        HorizontalAlignmentValues.Center,
                        true,
                        2
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
                        new Column("SheetName", "Team Leads", NumberFormats.General, HorizontalAlignmentValues.Left,
                            true, 0),
                        new Column("Age", "How Old?", NumberFormats.General, HorizontalAlignmentValues.Center, true, 1),
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
                        new Column("Registered", "Date", NumberFormats.TimeStamp124, HorizontalAlignmentValues.Left,
                            true, 3),
                        new Column("Member", "Member?", NumberFormats.General, HorizontalAlignmentValues.Center, true, 2),
                        new Column("SheetName", "Associate", NumberFormats.General, HorizontalAlignmentValues.Left, true,
                            0)
                    }
                ));

                /* -------------------------------------------------------------
                 * Setup and return the primary container for the child datasets
                 * Using Constructor Injection as well
                 * -----------------------------------------------------------*/
                return new Settings("Organization", childList);
            }
            catch (Exception ex)
            {
                Assert.Fail("Exception Fail: " + ex.Message);
            }
            return null;
        }
    }

    /// ===========================================================================================
    /// <summary>
    ///     Mock Data Generator
    /// </summary>
    /// ===========================================================================================
    public class CreateMockData
    {
        // Declarations
        private readonly DataSet _dataSet = new DataSet();

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     For multiple-tables there must always be a primary key -> foreign key relation,
        ///     otherwise the related child table will be skipped
        ///     ** Note: As an experiment comment one or more of the relations and test the result
        /// </summary>
        /// <returns>DataSet</returns>
        /// -----------------------------------------------------------------------------------------------
        public DataSet CreateDataSet()
        {
            try
            {
                _dataSet.Tables.Add(CreateDirectors());
                _dataSet.Tables.Add(CreateManagers());
                _dataSet.Tables.Add(CreateTeamLeads());
                _dataSet.Tables.Add(CreateAssociates());

                _dataSet.Relations.Add("FK_Managers_Directors",
                    _dataSet.Tables["Directors"].Columns["DID"],
                    _dataSet.Tables["Managers"].Columns["DID"]);

                _dataSet.Relations.Add("FK_TeamLeads_Managers",
                    _dataSet.Tables["Managers"].Columns["MID"],
                    _dataSet.Tables["TeamLeads"].Columns["MID"]);

                _dataSet.Relations.Add("FK_Associates_TeamLeads",
                    _dataSet.Tables["TeamLeads"].Columns["TLID"],
                    _dataSet.Tables["Associates"].Columns["TLID"]);

                return _dataSet;
            }
            catch (Exception ex)
            {
                Assert.Fail("Exception Fail: " + ex.Message);
            }
            return null;
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Create the Associates table
        /// </summary>
        /// <returns>DataTable</returns>
        /// -----------------------------------------------------------------------------------------------
        public DataTable CreateAssociates()
        {
            try
            {
                var table = new DataTable("Associates");
                table.Columns.Add("AID", typeof(int));
                table.Columns.Add("SheetName", typeof(string));
                table.Columns.Add("Age", typeof(int));
                table.Columns.Add("Income", typeof(double));
                table.Columns.Add("Member", typeof(bool));
                table.Columns.Add("Registered", typeof(DateTime));
                table.Columns.Add("TLID", typeof(int));

                table.PrimaryKey = new[] {table.Columns["AID"]};

                var newRow = table.NewRow();
                newRow["AID"] = 1;
                newRow["SheetName"] = "Dan";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["TLID"] = 3;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["AID"] = 2;
                newRow["SheetName"] = "Samuel L";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["TLID"] = 3;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["AID"] = 3;
                newRow["SheetName"] = "Samuel P";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["TLID"] = 3;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["AID"] = 4;
                newRow["SheetName"] = "Samuel D";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["TLID"] = 3;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["AID"] = 5;
                newRow["SheetName"] = "Kyle A";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["TLID"] = 6;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["AID"] = 6;
                newRow["SheetName"] = "Kyle B";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["TLID"] = 6;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["AID"] = 7;
                newRow["SheetName"] = "Kyle C";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["TLID"] = 6;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["AID"] = 8;
                newRow["SheetName"] = "Kyle D";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["TLID"] = 6;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["AID"] = 9;
                newRow["SheetName"] = "Kyle E";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["TLID"] = 6;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["AID"] = 10;
                newRow["SheetName"] = "Kyle F";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["TLID"] = 6;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["AID"] = 11;
                newRow["SheetName"] = "Kyle G";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["TLID"] = 6;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["AID"] = 12;
                newRow["SheetName"] = "Kyle H";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["TLID"] = 6;
                table.Rows.Add(newRow);

                return table;
            }
            catch (Exception ex)
            {
                Assert.Fail("Exception Fail: " + ex.Message);
            }
            return null;
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Create the Team Leads table
        /// </summary>
        /// <returns>DataTable</returns>
        /// -----------------------------------------------------------------------------------------------
        public DataTable CreateTeamLeads()
        {
            try
            {
                var table = new DataTable("TeamLeads");
                table.Columns.Add("TLID", typeof(int));
                table.Columns.Add("SheetName", typeof(string));
                table.Columns.Add("Age", typeof(int));
                table.Columns.Add("Income", typeof(double));
                table.Columns.Add("Member", typeof(bool));
                table.Columns.Add("Registered", typeof(DateTime));
                table.Columns.Add("MID", typeof(int));

                table.PrimaryKey = new[] {table.Columns["TLID"]};

                var newRow = table.NewRow();
                newRow["TLID"] = 1;
                newRow["SheetName"] = "Mary";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["MID"] = 2;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["TLID"] = 2;
                newRow["SheetName"] = "Peter";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["MID"] = 2;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["TLID"] = 3;
                newRow["SheetName"] = "Authur";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["MID"] = 3;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["TLID"] = 4;
                newRow["SheetName"] = "Willa";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["MID"] = 3;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["TLID"] = 5;
                newRow["SheetName"] = "Jack";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["MID"] = 4;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["TLID"] = 6;
                newRow["SheetName"] = "Ann";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["MID"] = 5;
                table.Rows.Add(newRow);

                return table;
            }
            catch (Exception ex)
            {
                Assert.Fail("Exception Fail: " + ex.Message);
            }
            return null;
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Create the Managers table
        /// </summary>
        /// <returns>DataTable</returns>
        /// -----------------------------------------------------------------------------------------------
        public DataTable CreateManagers()
        {
            try
            {
                var table = new DataTable("Managers");
                table.Columns.Add("MID", typeof(int));
                table.Columns.Add("SheetName", typeof(string));
                table.Columns.Add("Age", typeof(int));
                table.Columns.Add("Income", typeof(double));
                table.Columns.Add("Member", typeof(bool));
                table.Columns.Add("Registered", typeof(DateTime));
                table.Columns.Add("DID", typeof(int));

                table.PrimaryKey = new[] {table.Columns["MID"]};

                var newRow = table.NewRow();
                newRow["MID"] = 2;
                newRow["SheetName"] = "Sam";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["DID"] = 34;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["MID"] = 3;
                newRow["SheetName"] = "Andrew";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["DID"] = 34;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["MID"] = 4;
                newRow["SheetName"] = "Martha";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["DID"] = 72;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["MID"] = 5;
                newRow["SheetName"] = "Sonja";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["DID"] = 72;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["MID"] = 7;
                newRow["SheetName"] = "Joe";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["DID"] = 90;
                table.Rows.Add(newRow);

                return table;
            }
            catch (Exception ex)
            {
                Assert.Fail("Exception Fail: " + ex.Message);
            }
            return null;
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Create the Directors table
        /// </summary>
        /// <returns>DataTable</returns>
        /// -----------------------------------------------------------------------------------------------
        public DataTable CreateDirectors()
        {
            try
            {
                var table = new DataTable("Directors");
                table.Columns.Add("DID", typeof(int));
                table.Columns.Add("SheetName", typeof(string));
                table.Columns.Add("Age", typeof(int));
                table.Columns.Add("Income", typeof(double));
                table.Columns.Add("Member", typeof(bool));
                table.Columns.Add("Registered", typeof(DateTime));
                table.PrimaryKey = new[] {table.Columns["DID"]};

                var newRow = table.NewRow();
                newRow["DID"] = 15;
                newRow["SheetName"] = "Allen";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["DID"] = 34;
                newRow["SheetName"] = "Bill";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["DID"] = 72;
                newRow["SheetName"] = "Markus";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["DID"] = 90;
                newRow["SheetName"] = "Thomas";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                table.Rows.Add(newRow);

                return table;
            }
            catch (Exception ex)
            {
                Assert.Fail("Exception Fail: " + ex.Message);
            }
            return null;
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Destroy all the left over objects
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        ~CreateMockData()
        {
            try
            {
                _dataSet.Dispose();
            }
            catch (Exception ex)
            {
                Assert.Fail("Exception Fail: " + ex.Message);
            }
        }
    }
}