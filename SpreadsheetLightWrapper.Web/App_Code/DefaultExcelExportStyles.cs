using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLightWrapper.Core.style;
using SpreadsheetLightWrapper.Export.Models;
//using SpreadsheetLightWrapper.Properties;
using Color = System.Drawing.Color;
using Settings = SpreadsheetLightWrapper.Export.Models.Settings;

namespace SpreadsheetLightWrapper.Web
{
    /// ===========================================================================================
    /// <summary>
    ///     This static class sets up a set of Default Styles setting for the
    ///     Ups.Toolkit.SpreadsheetLight Excel Export class if no User-Defined Settings are setup
    ///     by the developer.
    /// </summary>
    /// ===========================================================================================
    public static class DefaultExcelExportStyles
    {
        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Sets up the default styling when user does not predefine styles
        ///     with a Settings configuration, using the property injection technique.
        /// </summary>
        /// <returns>Settings: Default Styling</returns>
        /// -----------------------------------------------------------------------------------------------
        public static Settings SetupDefaultStyles()
        {
            try
            {
                /* -------------------------------------------------------------
                * Setup primary container for the child datasets
                * -----------------------------------------------------------*/
                var settings = new Settings
                {
                    // Optional name
                    Name = "Default Settings Container"
                };

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
                baseColumnHeaderStyle.SetFont("Helvetica", 11);
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

                /* -------------------------------------------------------------
                 * Define and style base child settings.
                 * This Child will always be present, it represents the
                 * primary dataset for every export.
                 * -----------------------------------------------------------*/
                settings.ChildSettings.Add(new ChildSetting
                (
                    // Name (Optional)
                    name: "Default Base Child Settings",
                    // Set Overall Column Visibility
                    showColumnHeader: true,
                    // Column offset to the right
                    columnOffset: 0,
                    // Make the base column header row a little larger
                    // so it will stand out.  Value is in pixels
                    columnHeaderRowHeight: 25,
                    // Setup the style for Column Headers
                    columnHeaderStyle: baseColumnHeaderStyle,
                    // Row and Alternating Row Styles
                    // If set to false then the odd row style will be overall row style
                    showAlternatingRows: true,
                    // Setup the style for odd & even rows
                    oddRowStyle: oddRowStyle,
                    evenRowStyle: evenRowStyle,
                    // No User-Defined column headers
                    userDefinedColumns: null
                ));

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

                /* -------------------------------------------------------------
                 * Define and add the stylings for the first child, which is
                 * a child of the base data-set
                 * -----------------------------------------------------------*/
                settings.ChildSettings.Add(new ChildSetting
                (
                    "Default First Child Settings",
                    true,
                    null,
                    null,
                    firstColumnHeaderStyle,
                    true,
                    oddRowStyle,
                    evenRowStyle,
                    null
                ));

                /* -------------------------------------------------------------
                 * The second child column headers stylings will be derived
                 * from the first, change only what needs to be changed.
                 * -----------------------------------------------------------*/
                var secondColumnHeaderStyle = firstColumnHeaderStyle.Clone();
                secondColumnHeaderStyle.Fill.SetPattern(PatternValues.Solid, Color.CadetBlue, Color.White);
                secondColumnHeaderStyle.SetFontColor(Color.White);

                /* -------------------------------------------------------------
                 * Define and add the stylings for the second child, which is
                 * a child of the first data-set
                 * -----------------------------------------------------------*/
                settings.ChildSettings.Add(new ChildSetting
                (
                    "Default Second Child Settings",
                    true,
                    null,
                    null,
                    secondColumnHeaderStyle,
                    true,
                    oddRowStyle,
                    evenRowStyle,
                    null
                ));

                /* -------------------------------------------------------------
                 * The third child column headers stylings will be derived
                 * from the first, change only what needs to be changed.
                 * -----------------------------------------------------------*/
                var thirdColumnHeaderStyle = firstColumnHeaderStyle.Clone();
                thirdColumnHeaderStyle.Fill.SetPattern(PatternValues.Solid, Color.Aqua, Color.Black);
                thirdColumnHeaderStyle.SetFontColor(Color.Black);

                /* -------------------------------------------------------------
                 * Define and add the stylings for the third child, which is
                 * a child of the second data-set
                 * -----------------------------------------------------------*/
                settings.ChildSettings.Add(new ChildSetting
                (
                    "Default Third Child Settings",
                    true,
                    null,
                    null,
                    thirdColumnHeaderStyle,
                    true,
                    oddRowStyle,
                    evenRowStyle,
                    null
                ));

                /* -------------------------------------------------------------
                 * The forth child column headers stylings will be derived
                 * from the first, change only what needs to be changed.
                 * -----------------------------------------------------------*/
                var fourthColumnHeaderStyle = firstColumnHeaderStyle.Clone();
                fourthColumnHeaderStyle.Fill.SetPattern(PatternValues.Solid, Color.Chartreuse, Color.Black);
                fourthColumnHeaderStyle.SetFontColor(Color.Black);

                /* -------------------------------------------------------------
                 * Define and add the stylings for the fourth child, which is
                 * a child of the third data-set
                 * -----------------------------------------------------------*/
                settings.ChildSettings.Add(new ChildSetting
                (
                    "Default Fourth Child Settings",
                    true,
                    null,
                    null,
                    fourthColumnHeaderStyle,
                    true,
                    oddRowStyle,
                    evenRowStyle,
                    null
                ));

                /* -------------------------------------------------------------
                 * If five deep isn't enough let's add a sixth one.
                 * The fifth child column headers stylings will be derived
                 * from the first, change only what needs to be changed.
                 * -----------------------------------------------------------*/
                var fifthColumnHeaderStyle = firstColumnHeaderStyle.Clone();
                fifthColumnHeaderStyle.Fill.SetPattern(PatternValues.Solid, Color.BlueViolet, Color.Black);
                fifthColumnHeaderStyle.SetFontColor(Color.White);

                /* -------------------------------------------------------------
                 * Define and add the stylings for the fifth child, which is
                 * a child of the fourth data-set
                 * -----------------------------------------------------------*/
                settings.ChildSettings.Add(new ChildSetting
                (
                    "Default Fifth Child Settings",
                    true,
                    null,
                    null,
                    fifthColumnHeaderStyle,
                    true,
                    oddRowStyle,
                    evenRowStyle,
                    null
                ));
                
                return settings;
            }
            catch (Exception ex)
            {

            }
            return null;
        }
    }
}