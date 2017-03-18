using System;
using System.Collections.Generic;

namespace SpreadsheetLightWrapper.Export.Models
{
    /// ===========================================================================================
    /// <summary>
    ///     Allows the programmer to set custom properties
    ///     Acts as a container for the ChildSetting classes
    /// </summary>
    /// ===========================================================================================
    public class Settings
    {
        #region Properties

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Properties
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        public string Name { get; set; }

        public List<ChildSetting> ChildSettings { get; set; }

        #endregion Properties

        #region Constructors

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Overload 1: Base Constructor
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        public Settings()
        {
            try
            {
                Name = string.Empty;
                ChildSettings = new List<ChildSetting>();
            }
            catch (Exception ex)
            {
                //WebLogger.LogException(
                //    new Exception(
                //        "Ups.Toolkit.SpreadsheetLight.Export.Models.Settings.Contructor:Overload 1 -> " +
                //        ex.Message, ex),
                //    new Dictionary<string, string> {{"Settings", "Constructor:Overload 1"}});
            }
        }

        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Overload 2: Constructor - Set all properties
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        public Settings(
            string name,
            List<ChildSetting> childSettings)
        {
            try
            {
                Name = name ?? string.Empty;
                ChildSettings = childSettings ?? new List<ChildSetting>();
            }
            catch (Exception ex)
            {
                //WebLogger.LogException(
                //    new Exception(
                //        "Ups.Toolkit.SpreadsheetLight.Export.Models.Settings.Contructor:Overload 2 -> " +
                //        ex.Message, ex),
                //    new Dictionary<string, string> {{"Settings", "Constructor:Overload 2"}});
            }
        }

        #endregion Constructors
    }
}