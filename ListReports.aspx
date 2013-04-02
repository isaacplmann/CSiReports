<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ListReports.aspx.cs" Inherits="Reports_ListReports" %>
<%@ Register Assembly="CrystalDecisions.Web, Version=13.0.2000.0, Culture=neutral, PublicKeyToken=692fbea5521e1304" Namespace="CrystalDecisions.Web" TagPrefix="CR" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link type="text/css" rel="stylesheet" href="css/styles.css" />
    <link type="text/css" rel="stylesheet" href="css/normalize.css" />
</head>
<body>
    <form id="form1" runat="server">
        <div class="header">
                    <table style="width:100%;">
                        <tr>
                            <td style="text-align:right">
                                <a href="changepassword.aspx">Change Password</a>&nbsp;&nbsp;&nbsp;&nbsp;
                    <% if(HttpContext.Current.User == null || HttpContext.Current.User.Identity.Name.Length == 0) { %>
                    <a href="/Login.aspx">Log in</a>
                    <% } else { %>
                    <a href="/Logout.aspx">Log out</a>
                    <% } %>
                                &nbsp;&nbsp;&nbsp;&nbsp;
                            </td>
                        </tr>
                    </table>
        </div>
                    <div class="leftcolumn">
                    <img class="logo" alt="CSi Logo" src="images/csilogo.jpg" /><br />
                    <asp:Repeater runat="server" ID="leftmenu" OnItemCommand="LoadReport">
                        <ItemTemplate>
                            <div>
                                <%# Eval("Path").ToString().Length==0?"<h2>"+Eval("Name")+"</h2>":"" %>
                                <asp:LinkButton runat="server" Visible='<%# Eval("Path").ToString().Length>0 %>' CommandArgument='<%# Eval("Path") %>'><%# Eval("Name") %></asp:LinkButton>
                            </div>
                        </ItemTemplate>
                    </asp:Repeater>
                    </div>
        <div class="reportarea">
            <asp:Label runat="server" CssClass="introtext" ID="intro">Click on a report in the list on the left to display it.</asp:Label>
                    <CR:CrystalReportViewer ID="ReportViewer" runat="server" AutoDataBind="true" EnableDatabaseLogonPrompt="false"
                        EnableParameterPrompt="true" HasToggleParameterPanelButton="false"
                        Width="350px" Height="50px"/>
                    <CR:CrystalReportSource ID="CrystalReportSource1" runat="server" >
                        <Report FileName="c:\\Users\\fitpc\\documents\\visual studio 2012\\WebSites\\CSiReports\\VTAIGTrend.rpt">
                        </Report>
                    </CR:CrystalReportSource>
        </div>
    </form>
</body>
</html>
