<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="SpreadsheetLightWrapper.Web.Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Excel Export Tester</title>
    <link href="Content/bootstrap.css" rel="stylesheet"/>
    <link href="Content/Site.css" rel="stylesheet"/>
    <script src="Scripts/jquery-3.1.1.js"></script>
    <script src="Scripts/bootstrap.js"></script>
</head>
<body>
<form id="frmMain" runat="server">
    <div >
        <div class="jumbotron-mod">
            <div class="row">
                <div class="col-lg-10">
                    <p class="lead" style="font-size: x-large;">Export Examples with the Export Wrapper contained in SpreadsheetLightWrapper with the SpreadsheetLight Core</p>
                    <p class="lead">
                        <a target="_blank" href="http://spreadsheetlight.com/">SpreadsheetLight</a>
                    </p>
                    <p class="lead">
                        <a target="_blank" href="UsingtheExcelExportHelper.mht">Using the Excel Export Helper</a>
                    </p>
                </div>
            </div>
            <div class="row">
                <hr class="hr-mod"/>
            </div>
            <div class="row">
                <div class="col-lg-10">
                    <table class="nav-justified">
                        <tr>
                            <td>
                                <asp:Button ID="btnBasicRelatedGroupedDataSet" CssClass="btn btn-primary buttonSize" runat="server" Text="Basic Grouped Parent-Child Data" OnClick="btnBasicRelatedGroupedDataSet_Click"/>
                            </td>
                            <td style="width: 10px">&nbsp;</td>
                            <td class="commentSize">
                                There are no User-defined columns and DefaultSettings are used, and since there are no User-Defined columns the data is stringified with no numeric formatting.
                            </td>
                        </tr>
                        <tr>
                            <td colspan="3" style="height: 10px"></td>
                        </tr>
                        <tr>
                            <td>
                                <asp:Button ID="btnStyledRelatedGroupedDataSet" CssClass="btn btn-primary buttonSize" runat="server" Text="Styled Grouped Parent-Child Data" OnClick="btnStyledRelatedGroupedDataSet_Click"/>
                            </td>
                            <td style="">&nbsp;</td>
                            <td class="commentSize">
                                This example has User-defined columns and Customized Settings. The User-Defined columns set the column order, custom column names, visibility and data formatting.
                            </td>
                        </tr>
                        <tr>
                            <td colspan="3" style="height: 10px"></td>
                        </tr>
                        <tr>
                            <td>
                                <asp:Button ID="btnStyledRelatedGroupedDataSetToFile" CssClass="btn btn-primary buttonSize" runat="server" Text="Styled Grouped To File" OnClick="btnStyledRelatedGroupedDataSetToFile_Click"/>
                            </td>
                            <td style="">&nbsp;</td>
                            <td class="commentSize">
                                This example also has User-defined columns and Customized Settings; the User-Defined columns set the column order, custom column names, visibility and data formatting.
                                In this instance the output will be save to C:\SpreadsheetLightOutput.xls
                            </td>
                        </tr>
                        <tr>
                            <td colspan="3" style="height: 10px"></td>
                        </tr>
                        <tr>
                            <td>
                                <asp:Button ID="btnBasicUnrelatedUngroupedDataSet" CssClass="btn btn-primary buttonSize" runat="server" Text="Ungrouped No Parent-Child Relations" OnClick="btnUngroupedNoParentChildData_Click"/>
                            </td>
                            <td style="">&nbsp;</td>
                            <td class="commentSize">
                                It has four tables Directors, Managers, TeamLeads & Associates that are unrelated and each gets its own sheet in the Workbook
                            </td>
                        </tr>
                        <tr>
                            <td colspan="3" style="height: 10px"></td>
                        </tr>
                        <tr>
                            <td>
                                <asp:Button ID="butPartiallyRelatedGroupedDataSet" CssClass="btn btn-primary buttonSize" runat="server" Text="Partially Related 1 Parent-Child Relation" OnClick="btnPartiallyRelatedGroupedData_Click"/>
                            </td>
                            <td style="">&nbsp;</td>
                            <td class="commentSize">
                                This example has same four tables with two that are related, TeamLeads & Associates. The unrelated tables, Directors &amp; Managers, get their own sheets, the two related will be grouped on one sheet.
                            </td>
                        </tr>
                        <tr>
                            <td colspan="3" style="height: 10px"></td>
                        </tr>
                        <tr>
                            <td>
                                <asp:Button ID="btnPartiallyRelatedGroupedDataVer2" CssClass="btn btn-primary buttonSize" runat="server" Text="Partially Related 1 Parent-Child Relation" OnClick="btnPartiallyRelatedGroupedDataVer2_Click"/>
                            </td>
                            <td style="">&nbsp;</td>
                            <td class="commentSize">
                                This example has two that are related, Managers & TeamLeads. The unrelated tables, Directors &amp; Associates, get their own sheets, the two related will be grouped on one sheet.
                            </td>
                        </tr>
                        <tr>
                            <td colspan="3" style="height: 10px"></td>
                        </tr>
                        <tr>
                            <td>
                                <asp:Button ID="btnPartiallyRelatedGroupedDataVer3" CssClass="btn btn-primary buttonSize" runat="server" Text="Partially Related 2 Parent-Child Relations" OnClick="btnPartiallyRelatedGroupedDataVer3_Click"/>
                            </td>
                            <td style="">&nbsp;</td>
                            <td class="commentSize">
                                This example has three that are related, Directors, Managers & TeamLeads. The unrelated table Associates will get its own sheet, the three related will be grouped on one sheet.
                            </td>
                        </tr>
                    </table>
                </div>
            </div>
        </div>
    </div>
</form>
</body>
</html>