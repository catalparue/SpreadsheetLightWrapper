using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Web;

namespace Ups.Toolkit.SpreadsheetLight.Error
{
    public static class WebLogger
    {
        private static Logger.Logger _logger = new Logger.Logger();
        private static readonly string _sessionCustomPropsKeyName = "LoggerCustomProps";

        #region Public Static Methods

        public static void AddSavedCustomProperties(string key, string value)
        {
            AddSavedCustomProperties(new Dictionary<string, string> {{key, value}});
        }

        public static void AddSavedCustomProperties(Dictionary<string, string> props)
        {
            //Getting Current Saved Properties
            var savedProps = GetSavedCustomProperties();

            //Returning if there is nothing to add
            if ((props == null) && !props.Any())
                return;

            //Getting Http Context
            var context = HttpContext.Current;
            if ((context == null) && (context.Session != null))
                return;

            //Adding to saved Properties if any other wise add property to session
            if (savedProps.Any())
            {
                foreach (var pair in props.Where(x => !savedProps.ContainsKey(x.Key)))
                    savedProps.Add(pair.Key, pair.Value);
                context.Session[_sessionCustomPropsKeyName] = savedProps;
            }
            else
            {
                context.Session.Add(_sessionCustomPropsKeyName, props);
            }
        }

        public static void ClearSavedCustomProperties()
        {
            //Getting Current Properties
            var currentProps = GetSavedCustomProperties();

            if (!currentProps.Any())
                return;

            //Getting Http Context
            var context = HttpContext.Current;
            if (context == null)
                return;

            //Checking if Session and Key exist
            if ((context.Session != null) &&
                (context.Session[_sessionCustomPropsKeyName] != null))
                context.Session.Remove(_sessionCustomPropsKeyName);
        }

        public static Dictionary<string, string> GetSavedCustomProperties()
        {
            //Creating return object
            var ret = new Dictionary<string, string>();

            //Getting Http Context
            var context = HttpContext.Current;
            if (context == null)
                return ret;

            //Checking if Session and Key exist
            if ((context.Session != null) &&
                (context.Session[_sessionCustomPropsKeyName] != null))
            {
                //Casting session object
                ret = context.Session[_sessionCustomPropsKeyName] as Dictionary<string, string>;

                //Checking Return for object
                if (ret == null)
                    ret = new Dictionary<string, string>();

                //Removing Session Key if no items
                if (!ret.Any())
                    context.Session.Remove(_sessionCustomPropsKeyName);
            }

            //Returning dictionary
            return ret;
        }

        public static void LogError(string message, Dictionary<string, string> props = null, string controller = null,
            string action = null)
        {
            //Creating Logger, Logging and Clearing
            var logger = _createLogger(props);
            logger.LogError(message);
            logger = null;
        }

        public static void LogException(Exception ex, Dictionary<string, string> props = null, string controller = null,
            string action = null)
        {
            //Getting Inner Most Exception
            while (ex.InnerException != null)
                ex = ex.InnerException;

            //Creating Logger, Logging and Clearing
            var logger = _createLogger(props, ex);
            logger.LogException(ex);
            logger = null;
        }

        public static void LogInfo(string info, Dictionary<string, string> props = null, string controller = null,
            string action = null)
        {
            //Creating Logger, Logging and Clearing
            var logger = _createLogger(props);
            logger.LogInfoMessage(info);
            logger = null;
        }

        public static void LogWarning(string message, Dictionary<string, string> props = null, string controller = null,
            string action = null)
        {
            //Creating Logger, Logging and Clearing
            var logger = _createLogger(props);
            logger.LogWarningMessage(message);
            logger = null;
        }

        public static void SetSavedCustomProperties(string key, string value)
        {
            SetSavedCustomProperties(new Dictionary<string, string> {{key, value}});
        }

        public static void SetSavedCustomProperties(Dictionary<string, string> props)
        {
            //Clearing Current Custom Properties
            ClearSavedCustomProperties();

            //Checking if any properties to update
            if ((props == null) || !props.Any())
                return;

            //Adding Properties if Session
            var context = HttpContext.Current;
            if (context == null)
                return;
            context.Session.Add(_sessionCustomPropsKeyName, props);
        }

        public static void UpdateSavedCustomProperties(string key, string value)
        {
            UpdateSavedCustomProperties(new Dictionary<string, string> {{key, value}});
        }

        public static void UpdateSavedCustomProperties(Dictionary<string, string> props)
        {
            //Getting Current Saved Properties
            var savedProps = GetSavedCustomProperties();

            //Returning if there is nothing to add
            if ((props == null) && !props.Any())
                return;

            //Getting Http Context
            var context = HttpContext.Current;
            if ((context == null) && (context.Session != null))
                return;

            if (!savedProps.Any())
                foreach (var pair in props)
                    if (savedProps.ContainsKey(pair.Key))
                        savedProps[pair.Key] = pair.Value;
                    else
                        savedProps.Add(pair.Key, pair.Value);
            else
                context.Session.Add(_sessionCustomPropsKeyName, props);
        }

        #endregion Public Static Methods

        #region Private Static Methods

        private static Logger.Logger _createLogger(Dictionary<string, string> props, Exception ex = null)
        {
            //Creating Initial Logger
            var logger = new Logger.Logger();

            //Checking for Exception type
            if (ex != null)
                if (ex.GetType() == typeof(SqlException))
                {
                    var sqlEx = ex as SqlException;
                    if (!string.IsNullOrEmpty(sqlEx.Procedure.Trim()))
                        logger.CustomProperties.Add("StoredProcedure", sqlEx.Procedure);
                    logger.CustomProperties.Add("Server", sqlEx.Server);
                    if (sqlEx.LineNumber > 0)
                        logger.CustomProperties.Add("LineNumber", sqlEx.LineNumber.ToString());
                    if (!string.IsNullOrEmpty(sqlEx.State.ToString().Trim()))
                        logger.CustomProperties.Add("State", sqlEx.State.ToString().Trim());
                    if (sqlEx.Number > 0)
                        logger.CustomProperties.Add("Number", sqlEx.Number.ToString());
                }

            //Getting Http Context
            var context = HttpContext.Current;
            if (context != null)
            {
                //Getting User Information
                string companyAlias = null; //Comes from Passport External Authorzation, Workgroup SheetName used instead
                string loginId = null;

                //Getting HTTP Request
                var request = context.Request;
                if (request != null)
                {
                    //Getting Query Strings
                    if ((request.QueryString != null) &&
                        (request.QueryString.Count > 0))
                        logger.CustomProperties.Add("QueryString", _makeJsonString(request.QueryString));

                    //Getting Form Post Values
                    if ((request.Form != null) &&
                        (request.Form.Count > 0))
                        logger.CustomProperties.Add("Form", _makeJsonString(request.Form));
                } //if (request != null) {

                //Adding Session Items
                logger.CustomProperties.Add("SessionID", context.Session.SessionID);
                var sessionVars = new NameValueCollection();
                foreach (var key in context.Session.Keys)
                {
                    var obj = context.Session[key.ToString()];
                    if (obj != null)
                    {
                        var val = obj.ToString();
                        if (obj.GetType() == typeof(Array))
                            val += string.Format(" - ({0})", ((Array) obj).Length);
                        else if (obj.GetType() == typeof(IList))
                            val += string.Format(" - ({0})", ((IList) obj).Count);
                        else if (val.StartsWith("System.Collections.Generic.List`1"))
                            val += string.Format(" - ({0})", ((IList) obj).Count);
                        else if (obj.GetType() == typeof(IDictionary))
                            val += string.Format(" - ({0})", ((IDictionary) obj).Count);
                        else if (val.StartsWith("System.Collections.Generic.Dictionary`2"))
                            val += string.Format(" - ({0})", ((IDictionary) obj).Count);
                        sessionVars.Add(key.ToString(), val);

                        /*****************************
                        //Custom Dashboard Section as Login and workgroup is held 
                        //in the Session with specific key names
                        //*****************************/
                        //if (key.ToString() == "_ORIGINAL_EMPLOYEE_")
                        //{
                        //    THR_Employee orgEmployee = (THR_Employee)context.Session["_ORIGINAL_EMPLOYEE_"];
                        //    sessionVars.Add("Original Employee", orgEmployee.LastName + ", " + orgEmployee.FirstName);
                        //}
                        //else if (key.ToString() == "EmployeeInfo")
                        //{
                        //    THR_Employee employee = (THR_Employee)context.Session["EmployeeInfo"];
                        //    loginId = employee.RACFID;
                        //    companyAlias = employee.WorkGroupName;
                        //    sessionVars.Add("Employee Full SheetName", employee.LastName + ", " + employee.FirstName);
                        //}
                    }
                }
                logger.CustomProperties.Add("SessionVariables", _makeJsonString(sessionVars));

                //Adding User through Specific Custom Property
                //Getting Server/workstation user if not found thorugh previous session search
                if (string.IsNullOrEmpty(loginId) &&
                    (context.User != null) &&
                    (context.User.Identity != null) &&
                    !string.IsNullOrEmpty(context.User.Identity.Name))
                {
                    var name = context.User.Identity.Name;
                    if (name.Contains("\\"))
                    {
                        companyAlias = name.Substring(0, name.IndexOf("\\"));
                        loginId = name.Substring(name.IndexOf("\\") + 1);
                    }
                    else
                    {
                        loginId = name;
                    }
                }
                logger.CustomProperties.Add("LoginID", loginId);

                //Adding User through Specific Custom Property, found in session
                if (!string.IsNullOrEmpty(companyAlias))
                    logger.CustomProperties.Add("CompanyAlias", companyAlias);

                //Adding Client Information
                logger.CustomProperties.Add("ClientIP", request.UserHostAddress);
                logger.CustomProperties.Add("ClientBrowser",
                    string.Format("{0} {1}.{2}", request.Browser.Browser, request.Browser.MajorVersion,
                        request.Browser.MinorVersion));
                if (request.Browser.JScriptVersion != null)
                    logger.CustomProperties.Add("ClientJSVersion", request.Browser.JScriptVersion.ToString());
                logger.CustomProperties.Add("ClientPlatform", request.Browser.Platform);


                //Checking for Custom Properties that my be stored in session
                if ((context.Session != null) &&
                    (context.Session[_sessionCustomPropsKeyName] != null))
                {
                    //Getting Dictionary to add
                    var sessionProps = context.Session[_sessionCustomPropsKeyName] as Dictionary<string, string>;
                    if ((sessionProps != null) && sessionProps.Any())
                        foreach (var pair in sessionProps.Where(x => !logger.CustomProperties.ContainsKey(x.Key)))
                            logger.CustomProperties.Add(pair.Key, pair.Value);

                    //Removing Logger Custom Properties from Session
                    context.Session.Remove(_sessionCustomPropsKeyName);
                }
            } //if (context != null){

            //Adding Custom Properties from the parameter
            if ((props != null) && props.Any())
                foreach (var pair in props.Where(x => !logger.CustomProperties.ContainsKey(x.Key)))
                    logger.CustomProperties.Add(pair.Key, pair.Value);


            //Returning Logger
            return logger;
        }

        private static string _makeJsonString(NameValueCollection pairs)
        {
            //Checking Parameter
            if ((pairs == null) || (pairs.Count == 0))
                return "{}";

            //Creating String Builder
            var sb = new StringBuilder();

            //Start
            sb.Append("{");

            for (var i = 0; i < pairs.Count; i++)
            {
                var key = pairs.GetKey(i);

                //Removing double underscore keys that may be from ASP.NET
                if (key.StartsWith("__"))
                    continue;

                var value = string.Empty;
                try
                {
                    value = pairs.Get(i);
                }
                catch (HttpRequestValidationException)
                {
                    value = "Unsafe Request Form value.";
                }

                sb.Append(string.Format("\"{0}\":\"{1}\",",
                    key, value));
            }

            //End
            sb.Remove(sb.Length - 1, 1);
            sb.Append("}");

            //Returning SB contents
            return sb.ToString();
        }

        #endregion Private Static Methods
    }
}