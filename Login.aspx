<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Login.aspx.cs" Inherits="login" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>CSi Complete Login</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Login ID="Login1" runat="server" BackColor="#F7F7DE" BorderColor="#CCCC99" BorderStyle="Solid"
            BorderWidth="1px"
             DestinationPageUrl="ListReports.aspx" Font-Names="Verdana" Font-Size="10pt" PasswordRecoveryText="Forgot Password?"
            PasswordRecoveryUrl="recoverpassword.aspx">
        </asp:Login>
    </div>
    </form>
</body>
</html>
