using System;
using SpreadsheetLightWrapper.Core.misc;
using SpreadsheetLightWrapper.Core.worksheet;

namespace SpreadsheetLightWrapper.Core
{
    /// ===========================================================================================
    /// <summary>
    ///     Additional features added to the main SLDocument class
    /// </summary>
    /// ===========================================================================================
    public partial class SLDocument
    {
        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Custom function to add a single grouped row properties with OutlineLevel
        /// </summary>
        /// <param name="rowIndex">int</param>
        /// <param name="outLineLevel">int</param>
        /// -----------------------------------------------------------------------------------------------
        public void AddGroupedRow(int rowIndex, int outLineLevel)
        {
            try
            {
                if (rowIndex < SLConstants.RowLimit)
                {
                    var i = rowIndex;
                    SLRowProperties rp;
                    if (slws.RowProperties.ContainsKey(i))
                    {
                        rp = slws.RowProperties[i];
                        // Excel supports only 8 levels
                        if (rp.OutlineLevel < 8)
                        {
                            rp.OutlineLevel = (byte) outLineLevel;
                            if (rp.OutlineLevel > 0)
                                rp.Hidden = true;
                        }
                        slws.RowProperties[i] = rp.Clone();
                    }
                    else
                    {
                        rp = new SLRowProperties(SimpleTheme.ThemeRowHeight)
                        {
                            OutlineLevel = (byte) outLineLevel
                        };
                        slws.RowProperties[i] = rp.Clone();
                    }
                }
            }
            catch (Exception ex)
            {
                //WebLogger.LogException(
                //    new Exception(
                //        "Ups.Toolkit.SpreadsheetLight.Core.SLDocument.AddGroupedRow -> " + ex.Message, ex),
                //    new Dictionary<string, string> {{"SLDocument", "AddGroupedRow"}});
            }
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Collapses all the rows, called at the tail end of a grouping
        /// </summary>
        /// <param name="rowIndex">int</param>
        /// -----------------------------------------------------------------------------------------------
        public void CollapseAllRows(int rowIndex)
        {
            try
            {
                if (rowIndex < 1 || rowIndex > SLConstants.RowLimit) return;

                // the following algorithm is not guaranteed to work in all cases.
                // The data is sort of loosely linked together with no guarantee that they
                // all make sense together. If you use Excel, then the internal data is sort of
                // guaranteed to make sense together, but anyone can make an Open XML spreadsheet
                // with possibly invalid-looking data. Maybe Excel will accept it, maybe not.

                SLRowProperties rp;
                byte byCurrentOutlineLevel = 0;
                if (slws.RowProperties.ContainsKey(rowIndex))
                {
                    rp = slws.RowProperties[rowIndex];
                    byCurrentOutlineLevel = rp.OutlineLevel;
                }

                var bFound = false;
                int i;

                for (i = rowIndex - 1; i >= 1; --i)
                    if (slws.RowProperties.ContainsKey(i))
                    {
                        rp = slws.RowProperties[i];
                        if (rp.OutlineLevel > byCurrentOutlineLevel)
                        {
                            bFound = true;
                            rp.Hidden = true;
                            slws.RowProperties[i] = rp.Clone();
                        }
                        else
                        {
                            break;
                        }
                    }
                    else
                    {
                        break;
                    }

                if (bFound)
                    if (slws.RowProperties.ContainsKey(rowIndex))
                    {
                        rp = slws.RowProperties[rowIndex];
                        rp.Collapsed = true;
                        slws.RowProperties[rowIndex] = rp.Clone();
                    }
                    else
                    {
                        rp = new SLRowProperties(SimpleTheme.ThemeRowHeight) {Collapsed = true};
                        slws.RowProperties[rowIndex] = rp.Clone();
                    }
            }
            catch (Exception ex)
            {
                //WebLogger.LogException(
                //    new Exception(
                //        "Ups.Toolkit.SpreadsheetLight.Core.SLDocument.CollapseAllRows -> " + ex.Message, ex),
                //    new Dictionary<string, string> {{"SLDocument", "CollapseAllRows"}});
            }
        }
    }
}

/* Archived
// -----------------------------------------------------------------------------------------------
// <summary>
//     Collapses all the rows
// </summary>
// <param name="startRowIndex">int</param>
// <param name="endRowIndex">int</param>
// -----------------------------------------------------------------------------------------------
//public void CollapseAllRows(int startRowIndex, int endRowIndex)
//{
//    try
//    {
//        if (startRowIndex < 1 || startRowIndex > SLConstants.RowLimit) return;
//        if (endRowIndex < 1 || endRowIndex > SLConstants.RowLimit) return;

//        int rowIndex;
//        for (rowIndex = startRowIndex; rowIndex <= endRowIndex; rowIndex++)
//        {
//            SLRowProperties rp;
//            byte byCurrentOutlineLevel = 0;
//            if (slws.RowProperties.ContainsKey(rowIndex))
//            {
//                rp = slws.RowProperties[rowIndex];
//                byCurrentOutlineLevel = rp.OutlineLevel;
//            }

//            var bFound = false;
//            int i;

//            for (i = rowIndex - 1; i >= 1; --i)
//                if (slws.RowProperties.ContainsKey(i))
//                {
//                    rp = slws.RowProperties[i];
//                    if (rp.OutlineLevel > byCurrentOutlineLevel)
//                    {
//                        bFound = true;
//                        rp.Hidden = true;
//                        slws.RowProperties[i] = rp.Clone();
//                    }
//                    else
//                    {
//                        break;
//                    }
//                }
//                else
//                {
//                    break;
//                }

//            if (bFound)
//                if (slws.RowProperties.ContainsKey(rowIndex))
//                {
//                    rp = slws.RowProperties[rowIndex];
//                    rp.Collapsed = true;
//                    slws.RowProperties[rowIndex] = rp.Clone();
//                }
//                else
//                {
//                    rp = new SLRowProperties(SimpleTheme.ThemeRowHeight) {Collapsed = true};
//                    slws.RowProperties[rowIndex] = rp.Clone();
//                }
//        }
//    }
//    catch (Exception ex)
//    {
//        WebLogger.LogException(
//            new Exception(
//                "Ups.Toolkit.SpreadsheetLight.Core.SLDocument.CollapseAllRows -> " + ex.Message, ex),
//            new Dictionary<string, string> {{"SLDocument", "CollapseAllRows"}});
//    }
//}
*/