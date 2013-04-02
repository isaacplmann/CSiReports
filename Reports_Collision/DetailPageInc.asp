<%@ Language=VBScript %>
<%Response.Buffer=True%>
<html>
<head>
<title>Survey Detail</title>
<script language="Javascript">
function openZoom1(url, name) {
  popupWin = window.open(url, name,
"ScrollBars = yes, resizable=no,location=no,toolbar=no,top=6,left=155,width=750,height=470");
}
<!--
if (window.focus) {
  self.focus();
}
//-->
</script>
</head>
<body><font face=arial>
<p><font face=Arial>
<table width=100% border=0>
	<tr>
		<td>
		<img style="WIDTH: 278px; HEIGHT: 77px" height=44 hspace=1 src="http://www.csicomplete.com/images/logoreport.jpg" vspace=1 border=0>
		</td>
		<td align="center">
		<img style="WIDTH: 50px; HEIGHT: 30px" height=44 hspace=1 src="http://www.csicomplete.com/images/email.jpg" vspace=1 border=0>
		</td>
	</tr>
</table>
</font></P>
<%
Response.Flush
dim ConnString
dim jobid
dim shopname
dim strRolledOver

ConnString= session("ConString")
jobid = Request.QueryString("jobid")
'shopname = Request.QueryString("shopname")
iShopID = Request.QueryString("ShopID")

set cnList=server.CreateObject("adodb.connection")
cnlist.Open ConnString


strdetail = "Select name from tblCollisionLogons where shopid=" & ishopid
set rsdetail=server.createobject("adodb.recordset")
rsDetail.open strDetail,cnlist

if not rsdetail.eof then
	strshop=rsdetail.fields("name")
end if
rsdetail.close


strDetail = "select * from tblCollisionDetailAll inner join tblCollisionDetail on tblCollisionDetailAll.JobID=tblCollisionDetail.JobID where tblCollisionDetail.JobID=" & jobid & " "
set rsdetail=server.createobject("adodb.recordset")
rsDetail.open strDetail,cnlist
if rsdetail.eof then
	Response.Write "There is no such record." & "<br>"
	Response.Write "<br>"
else 
	response.write "<b>" & strshop & "</b><br>"
	
	if rsdetail.Fields("rolledover") = -1 then
		strRolledOver="YES"
	elseif rsdetail.Fields("rolledover")=0 then
		strRolledOver="NO"
	else
		strRolledOver="N/A"
	end if
	
%>

<table cellspacing=1 cellpadding=1 width=100% border=0>
<tr valign=top>
<td width="20%">Name:</td>
<td><b><font color=blue><%Response.write rsdetail.Fields("FirstName") & " " & rsdetail.Fields("LastName")%></b></FONT></td>
</tr>
<tr valign=top>
<td><font size=2>RO: </td>
<td><font size=2><%Response.write rsdetail.Fields("RO")%></font></td>
</tr>
<tr valign=top>
<td><font size=2>Home: </td>
<td><font size=2><%Response.write rsdetail.Fields("HomePhone")%></font></td>
</tr>
<tr valign=top>
<td><font size=2>Work: </td>
<td><font size=2><%Response.write rsdetail.Fields("WorkPhone")%></font></td>
</tr>
<tr valign=top>
<td><font size=2>RO Entered: </td>
<td><font size=2><%Response.write rsdetail.Fields("ROEntered")%></font></td>
</tr>
<tr valign=top>
<td><font size=2>Connections Made: </td>
<td><font size=2><%Response.write rsdetail.Fields("Calls")%></font></td>
</tr>
<tr valign=top>
<td><font size=2>Last Connected Call: </td>
<td><font size=2><%Response.write rsdetail.Fields("LastCall")%></font></td>
</tr>
<tr valign=top>
<td><font size=2>Call Notes: </td>
<td><font size=2><%Response.write rsdetail.Fields("CallNote")%></font></td>
</tr>
<tr valign=top>
<td colspan=2><font size=2>RO transferred from previous month: <%Response.write strRolledover%></font></td>
</tr>

<%
Response.Flush%>
</table>
<%
end if
%>

<p>
<center><strong><em><font size=1>
Information compiled by CSi Complete, a national provider of CSI and call center services to business.</br>
For more information please call (800) 343-0641. Copyright 2003 CSi Complete</font></B></I>
</font></em></strong></center>
</body></html>
