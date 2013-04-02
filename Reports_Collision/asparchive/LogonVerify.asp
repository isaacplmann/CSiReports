<%@ Language=VBScript %>
<%
session("ConString")= "Provider=SQLOLEDB;User ID=csiadmin;Password=admin;Persist Security Info=True;Initial Catalog=csi_data;Data Source=216.29.40.156"

set cn=server.CreateObject("adodb.connection")
cn.Open session("constring")

set rs = server.CreateObject("adodb.recordset")
strSQL = "SELECT L.CollisionID, L.Logon, L.Password, L.Name, L.ShopID, L.ParentID, L.RegionID, " & _
	" L.Type, L.SurveyTypeID FROM tblCollisionLogons L " & _
	" where L.logon = '" & request.form("txtName") & "' and L.Active=-1"
rs.Open strSQL,cn
if not rs.Eof then
	if Request.Form("txtPwd") = rs.Fields("password") then
		session("id") =rs.Fields("CollisionID")
		session("sessionname") = rs.Fields("Name")
		session("sessionparentid") = rs.Fields("parentid")
		session("sessionregionid")=rs.Fields("regionid")
		session("clienttype")=rs.Fields("Type")
		session("shopid")=rs.Fields("shopid")
		session("surveytypeid")=rs.Fields("surveytypeid")
		Response.Redirect("HomeCSi.asp")
	else
		Response.Redirect("../reports/LogonBad.htm")
	end if
else
	Response.Redirect("../reports/LogonBad.htm")
end if
rs.Close
cn.Close
%><HTML><HEAD><META content="MSHTML 6.00.2713.1100" name=GENERATOR></HEAD>
<BODY>LOGON VERIFY PAGE</BODY></HTML>