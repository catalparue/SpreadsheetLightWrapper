using System;
using System.Collections.Generic;
using System.Data;

namespace SpreadsheetLightWrapper.Web.Mocks
{
    public class Directors
    {
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
                table.Columns.Add("Name", typeof(string));
                table.Columns.Add("Age", typeof(int));
                table.Columns.Add("Income", typeof(double));
                table.Columns.Add("Member", typeof(bool));
                table.Columns.Add("Registered", typeof(DateTime));
                table.PrimaryKey = new[] {table.Columns["DID"]};

                var newRow = table.NewRow();
                newRow["DID"] = 15;
                newRow["Name"] = "Allen";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["DID"] = 34;
                newRow["Name"] = "Bill";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["DID"] = 72;
                newRow["Name"] = "Markus";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["DID"] = 90;
                newRow["Name"] = "Thomas";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                table.Rows.Add(newRow);

                return table;
            }
            catch (Exception ex)
            {

            }
            return null;
        }
    }
}