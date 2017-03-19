using System;
using System.Data;

namespace SpreadsheetLightWrapper.Web.Mocks
{
    public class TeamLeads
    {
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
                table.Columns.Add("Name", typeof(string));
                table.Columns.Add("Age", typeof(int));
                table.Columns.Add("Income", typeof(double));
                table.Columns.Add("Member", typeof(bool));
                table.Columns.Add("Registered", typeof(DateTime));
                table.Columns.Add("MID", typeof(int));

                table.PrimaryKey = new[] {table.Columns["TLID"]};

                var newRow = table.NewRow();
                newRow["TLID"] = 1;
                newRow["Name"] = "Mary";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["MID"] = 2;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["TLID"] = 2;
                newRow["Name"] = "Peter";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["MID"] = 2;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["TLID"] = 3;
                newRow["Name"] = "Authur";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["MID"] = 3;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["TLID"] = 4;
                newRow["Name"] = "Willa";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["MID"] = 3;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["TLID"] = 5;
                newRow["Name"] = "Jack";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["MID"] = 4;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["TLID"] = 6;
                newRow["Name"] = "Ann";
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
            }
            return null;
        }
    }
}