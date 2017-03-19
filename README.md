# SpreadsheetLightWrapper
Wrapper for SpreadsheetLight to more easily facilitate it use
Utility intended to easily aid developers setting up routines that will export data to Excel without the need of Microsoft Office.  The core library is the source code of Vincent Tan Wai Lip's SpreadsheetLight Version 3.4 (27 July 2013) http://spreadsheetlight.com/


Initializing the Exporter

The “Exporter” has static components make it very easy to use. The helper is flexible in that it will place separate DataTables with no parent-child relation on separate Sheets within the same Workbook, but if there is a relation then the tables will be grouped on the same Sheet. For example, if there are four tables in the DataSet, and two are related and two are not, then the two unrelated tables will appear on separate sheets, while the two related tables will grouped on the same Sheet. The helper is also flexible in area of styling; you can go all out and really customize the output with User-Defined Columns, and Styling for all sections, or you can just call the base “OutputWorkbook” function and it will just export the data with default styling. All of the static “OutputWorkbook” function overloads output bytes that will be used by webpage with a “Response” Header in this manner: First create the Data: In this instance mock data is being used, but you will create a DataTable or DataSet from the output of procedures that are accessed from the in-house DAL

View the "Using Excel Export Helper":
https://github.com/rdonalson/SpreadsheetLightWrapper/blob/master/Using%20the%20Excel%20Export%20Helper.pdf
