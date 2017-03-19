using System;
using System.Web;
using log4net;
using log4net.Config;

[assembly: XmlConfigurator(ConfigFile = "log4net.config", Watch = true)]

namespace SpreadsheetLightWrapper.Web
{
    public class Global : HttpApplication
    {
        private static readonly ILog Log = LogManager.GetLogger(typeof(Global));

        protected void Application_Start(object sender, EventArgs e)
        {
            /* Diagnostic */
            //Log.Info("Startup application.");
        }

        protected void Session_Start(object sender, EventArgs e)
        {
            /* Diagnostic */
            //Log.Info("Session Startup.");
        }

        protected void Application_BeginRequest(object sender, EventArgs e)
        {
        }

        protected void Application_AuthenticateRequest(object sender, EventArgs e)
        {
        }

        protected void Application_Error(object sender, EventArgs e)
        {
            // Get the exception object.
            var ex = Server.GetLastError();

            // Handle HTTP errors
            if (ex.GetType() == typeof(HttpException))
            {
                // The Complete Error Handling Example generates
                // some errors using URLs with "NoCatch" in them;
                // ignore these here to simulate what would happen
                // if a global.asax handler were not implemented.
                if (ex.Message.Contains("NoCatch") || ex.Message.Contains("maxUrlLength"))
                    return;

                //Redirect HTTP errors to HttpError page
                Server.Transfer("HttpErrorPage.aspx");
            }
            Log.Error("SpreadsheetLightWrapper.Web.Global.Application_Error -> " + ex.Message + ": " + ex);
        }

        protected void Session_End(object sender, EventArgs e)
        {
            /* Diagnostic */
            //Log.Info("Session End.");
        }

        protected void Application_End(object sender, EventArgs e)
        {
            /* Diagnostic */
            //Log.Info("Shutdown application.");
        }
    }
}