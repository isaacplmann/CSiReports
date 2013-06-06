<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ListReports.aspx.cs" Inherits="Reports_ListReports" %>
<%@ Register Assembly="CrystalDecisions.Web, Version=13.0.2000.0, Culture=neutral, PublicKeyToken=692fbea5521e1304" Namespace="CrystalDecisions.Web" TagPrefix="CR" %>
<%
if(HttpContext.Current.User == null || HttpContext.Current.User.Identity.Name.Length == 0) {
    Response.Redirect("/Login.aspx");
}
%>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link type="text/css" rel="stylesheet" href="css/normalize.css" />
    <link type="text/css" rel="stylesheet" href="css/styles.css" />
    <link type="text/css" rel="stylesheet" href="css/chosen/chosen.css" />
</head>
<body>
    <form id="form1" runat="server">
        <div class="header">
            <div class="nav account">
                <ul>
                    <li>
                        <span class="welcome">welcome, <asp:Label runat="server" ID="username"></asp:Label></span>
                    </li>
                    <li><a href="/changepassword.aspx">account</a></li>
                    <li><a href="/support.aspx">support</a></li>
                    <li><a href="/Logout.aspx">log out</a></li>
                </ul>
                <ul>
                    <li><a href="/help.aspx">help</a></li>
                </ul>
            </div>
            <img class="logo" alt="CSi Logo" src="images/csilogo.jpg" />
        </div>
        <div class="nav primary">
            <ul>
                <li id="DashboardItem" runat="server" class="dashboarditem isActive"><asp:LinkButton ID="DashboardLink" runat="server" OnCommand="ShowDashboard">Dashboard</asp:LinkButton></li>
                <li id="ExecutiveItem" runat="server" Visible="false" class="executiveitem"><asp:LinkButton ID="ExecutiveLink" runat="server" OnCommand="ChangeShopList" CommandName="Executive">Executive Report</asp:LinkButton></li>
                <li id="ShopListItem" runat="server" class="shoplistitem">
                    <div id="MultipleShops" runat="server">
                        <asp:Label Text="Shop Report: " AssociatedControlID="ShopList" runat="server" />
                        <asp:DropDownList ID="ShopList" runat="server" CssClass="chzn-select" data-placeholder="Choose shop..."
                            DataTextField="Name" DataValueField="Path" OnSelectedIndexChanged="ShopList_SelectedIndexChanged" SelectionMode="Single"
                            AutoPostBack="true">
                        </asp:DropDownList>
                    </div>
                    <asp:LinkButton ID="ShopLink" runat="server" Visible="false" OnCommand="ChangeShopList" CommandName="Shop">Shop Report</asp:LinkButton>
                </li>
                <li id="SurveyItem" runat="server" Visible="false" class="surveyitem">
                    <asp:HyperLink ID="SurveyLink" runat="server" Target="_blank" >Survey</asp:HyperLink>
                </li>
            </ul>
        </div>
        <div class="nav secondary">
            <ul>
                <asp:Repeater runat="server" ID="ReportList" OnItemCommand="LoadReport">
                    <ItemTemplate>
                        <li>
                            <asp:LinkButton ID="ReportLink" runat="server" Visible='<%# Eval("Path").ToString().Length>0 %>' CommandArgument='<%# Eval("Path") %>'><%# Eval("Name") %></asp:LinkButton>
                        </li>
                    </ItemTemplate>
                </asp:Repeater>
            </ul>
        </div>

        <div class="reportarea">
            <asp:Label runat="server" CssClass="introtext" ID="intro">Click on a report in the list on the left to display it.</asp:Label>
                    <CR:CrystalReportViewer ID="ReportViewer" runat="server" AutoDataBind="true" EnableDatabaseLogonPrompt="false"
                        EnableParameterPrompt="false" HasToggleParameterPanelButton="false" HasCrystalLogo="False"
                        Width="100%" Height="50px"/>
                    <CR:CrystalReportSource ID="CrystalReportSource1" runat="server" >
                        <Report FileName="VTAIGTrend.rpt">
                        </Report>
                    </CR:CrystalReportSource>
        </div>
    </form>
    <script src="js/jquery.min.js"></script>
    <script src="js/chosen.jquery.min.js" type="text/javascript"></script>
    <script type="text/javascript">
        $(".chzn-select").chosen({
            disable_search_threshold: 10
        });
    </script>
</body>
</html>
