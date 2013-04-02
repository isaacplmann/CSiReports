<%@ Language=VBScript %>
<%
if Request.Form("txtName")="" then
	Response.Redirect("LogonInvalidTest.htm")
end if
if Request.Form("txtPwd")="" then
	Response.Redirect("LogonInvalidTest.htm")
end if

session("ConString")= "Provider=SQLOLEDB;User ID=csiadmin;Password=admin;Persist Security Info=True;Initial Catalog=csi_data;Data Source=216.28.245.12"
set cn=server.CreateObject("adodb.connection")
cn.Open session("constring")
set rs = server.CreateObject("adodb.recordset")
strSQL = "select * from tblCollisionLogons where logon = '" & Request.Form("txtName") & "'"

Response.Write	"Step 1"

rs.Open strSQL,cn
if not rs.Eof then
	'they have a logon name
	'check the id to make sure they are who they are logged in as
	set rsLogon = server.CreateObject("adodb.recordset")
	strLogon = "Select CollisionID, Logon, Password from tblCollisionLogons where logon = '" & Request.Form("txtName") & "' and CollisionID=" & session("id")
	rslogon.Open strlogon, cn
		'if it's in the web_logons table
		if not rslogon.EOF then
			'make sure the new passwords are equal
			if Request.Form("txtPwd") = Request.Form("txtPwdConf") then 
			'update the passwords
				strUpdateL = "Update tblCollisionLogons set password='" & Request.Form("txtPwd") & "' where collisionid = " & session("id")
				cn.Execute strupdateL
				Response.Redirect("LogonSuccessTest.htm")
			else
			Response.Redirect("LogonInvalidTest.htm")
			end if						
		end if
		'if it wasn't in the first table
		'set rsAddl = server.CreateObject("adodb.recordset")
		'strAddl = "Select LogonID, UserName, Password from tbl_Addl_Logons where username = '" & Request.Form("txtName") & "' and LogonID=" & session("id")
		'rsaddl.Open strAddl, cn
		'if not rsaddl.EOF then
		'make sure the new passwords are equal
			'if Request.Form("txtPwd") = Request.Form("txtPwdConf") then 
			'update the passwords
				'strUpdateA = "Update tbl_Addl_Logons set password='" & Request.Form("txtPwd") & "' where username = '" & Request.Form("txtName") & "'and LogonID = " & session("id") 
				'cn.Execute strupdateA
				'Response.Redirect("LogonSuccess.htm")
			'else
			'Response.Redirect("LogonInvalid.htm")
			'end if						
		'end if
else
	Response.Redirect("LogonInvalidTest.htm")
end if
%><META content="MSHTML 6.00.2713.1100" name=GENERATOR></HEAD>
<BODY>PASSWORD VERIFY PAGE</BODY></HTML>




