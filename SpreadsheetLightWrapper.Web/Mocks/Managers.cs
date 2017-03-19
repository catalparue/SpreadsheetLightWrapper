using System;
using System.Collections.Generic;
using System.Data;

namespace SpreadsheetLightWrapper.Web.Mocks
{
    public class Managers
    {
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
                table.Columns.Add("Name", typeof(string));
                table.Columns.Add("Age", typeof(int));
                table.Columns.Add("Income", typeof(double));
                table.Columns.Add("Member", typeof(bool));
                table.Columns.Add("Registered", typeof(DateTime));
                table.Columns.Add("DID", typeof(int));

                table.PrimaryKey = new[] {table.Columns["MID"]};

                var newRow = table.NewRow();
                newRow["MID"] = 2;
                newRow["Name"] = "Sam";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["DID"] = 34;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["MID"] = 3;
                newRow["Name"] = "Andrew";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["DID"] = 34;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["MID"] = 4;
                newRow["Name"] = "Martha";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["DID"] = 72;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["MID"] = 5;
                newRow["Name"] = "Sonja";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["DID"] = 72;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["MID"] = 7;
                newRow["Name"] = "Joe";
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

            }
            return null;
        }
    }
}