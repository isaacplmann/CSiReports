<%@ Page Language="C#" AutoEventWireup="true" CodeFile="MyProfile.aspx.cs" Inherits="MyProfile" %>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>My Profile Page</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <h1>
            Profile Settings for: <asp:LoginName ID="LoginName1" runat="server" />
        </h1>
        
        <table>
            <tr>
                <td>Shop ID:</td>
                <td><asp:Label ID="ShopID" runat="server" Text="Label"></asp:Label></td>
            </tr>        
        
             <tr>
                <td valign="top">Roles:</td>
                <td><asp:ListBox ID="RoleList" runat="server" /></td>
            </tr>       
        
        </table>
        
    </div>
    </form>
</body>
</html>
