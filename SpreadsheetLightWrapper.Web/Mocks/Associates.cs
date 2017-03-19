using System;
using System.Data;

namespace SpreadsheetLightWrapper.Web.Mocks
{
    public class Associates
    {
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
                table.Columns.Add("Name", typeof(string));
                table.Columns.Add("Age", typeof(int));
                table.Columns.Add("Income", typeof(double));
                table.Columns.Add("Member", typeof(bool));
                table.Columns.Add("Registered", typeof(DateTime));
                table.Columns.Add("TLID", typeof(int));

                table.PrimaryKey = new[] {table.Columns["AID"]};

                var newRow = table.NewRow();
                newRow["AID"] = 1;
                newRow["Name"] = "Dan";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["TLID"] = 3;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["AID"] = 2;
                newRow["Name"] = "Samuel L";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["TLID"] = 3;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["AID"] = 3;
                newRow["Name"] = "Samuel P";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["TLID"] = 3;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["AID"] = 4;
                newRow["Name"] = "Samuel D";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["TLID"] = 3;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["AID"] = 5;
                newRow["Name"] = "Kyle A";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["TLID"] = 6;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["AID"] = 6;
                newRow["Name"] = "Kyle B";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["TLID"] = 6;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["AID"] = 7;
                newRow["Name"] = "Kyle C";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["TLID"] = 6;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["AID"] = 8;
                newRow["Name"] = "Kyle D";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["TLID"] = 6;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["AID"] = 9;
                newRow["Name"] = "Kyle E";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["TLID"] = 6;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["AID"] = 10;
                newRow["Name"] = "Kyle F";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["TLID"] = 6;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["AID"] = 11;
                newRow["Name"] = "Kyle G";
                newRow["Age"] = 30;
                newRow["Income"] = 100000;
                newRow["Member"] = true;
                newRow["Registered"] = DateTime.Now;
                newRow["TLID"] = 6;
                table.Rows.Add(newRow);

                newRow = table.NewRow();
                newRow["AID"] = 12;
                newRow["Name"] = "Kyle H";
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
            }
            return null;
        }
    }
}