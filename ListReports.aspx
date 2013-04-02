<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ListReports.aspx.cs" Inherits="Reports_ListReports" %>
<%@ Register Assembly="CrystalDecisions.Web, Version=13.0.2000.0, Culture=neutral, PublicKeyToken=692fbea5521e1304" Namespace="CrystalDecisions.Web" TagPrefix="CR" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <table>
            <tr>
                <td colspan="2">
                    Header content...
                    <% if(HttpContext.Current.User == null || HttpContext.Current.User.Identity.Name.Length == 0) { %>
                    <a href="/Login.aspx">Log in</a>
                    <% } else { %>
                    <a href="/Logout.aspx">Log out</a>
                    <% } %>
                </td>
            </tr>
            <tr>
                <td style="vertical-align:top">
                    <asp:Repeater runat="server" ID="leftmenu" OnItemCommand="LoadReport">
                        <ItemTemplate>
                            <div>
                                <%# Eval("Path").ToString().Length==0?"<h2>"+Eval("Name")+"</h2>":"" %>
                                <asp:LinkButton runat="server" Visible='<%# Eval("Path").ToString().Length>0 %>' CommandArgument='<%# Eval("Path") %>'><%# Eval("Name") %></asp:LinkButton>
                            </div>
                        </ItemTemplate>
                    </asp:Repeater>
                </td>
                <td>
                    <asp:Label runat="server" ID="intro" />
                    <CR:CrystalReportViewer ID="ReportViewer" runat="server" AutoDataBind="true" EnableDatabaseLogonPrompt="false"
                        EnableParameterPrompt="true" HasToggleParameterPanelButton="false"
                        Width="350px" Height="50px"/>
                    <CR:CrystalReportSource ID="CrystalReportSource1" runat="server" >
                        <Report FileName="c:\\Users\\fitpc\\documents\\visual studio 2012\\WebSites\\CSiReports\\VTAIGTrend.rpt">
                        </Report>
                    </CR:CrystalReportSource>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
