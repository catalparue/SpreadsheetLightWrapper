using System;
using System.Collections.Generic;
using System.Reflection;
using log4net;

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
        /// -----------------------------------------------------------------------------------------------
        /// <summary>
        ///     Internal Members
        /// </summary>
        /// -----------------------------------------------------------------------------------------------
        private static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

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
                Log.Error("SpreadsheetLightWrapper.Export.Models.Settings.Contructor:Overload 1 -> " + ex.Message + ": " +
                          ex);
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
                Log.Error("SpreadsheetLightWrapper.Export.Models.Settings.Contructor:Overload 2 -> " + ex.Message + ": " +
                          ex);
            }
        }

        #endregion Constructors
    }
}