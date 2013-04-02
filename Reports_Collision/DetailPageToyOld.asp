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
<font face=Arial><a href="http://www.csicomplete.com">
<table width=100% border=0>
	<tr>
		<td>
		<img style="WIDTH: 278px; HEIGHT: 77px" height=44 hspace=1 src="http://209.115.48.235/images/logoreport.jpg" vspace=1 border=0>
		</td>
		<td align="center">
		<img style="WIDTH: 50px; HEIGHT: 30px" height=44 hspace=1 src="http://209.115.48.235/images/email.jpg" vspace=1 border=0>
		</td>
	</tr>
</table>

</a></font><br>
<%
Response.Flush
dim ConnString
dim jobid
dim shopname
dim strClaimQ
dim strRecInsQ
dim strRolledOver

ConnString= "Provider=SQLOLEDB;User ID=csiadmin;Password=admin;Persist Security Info=True;Initial Catalog=csi_data;Data Source=216.28.245.12"
jobid = Request.QueryString("jobid")
'shopname = Request.QueryString("shopname")
ishopid=Request.QueryString("shopid")

set cnList=server.CreateObject("adodb.connection")
cnlist.Open ConnString

strdetail = "Select name from tblCollisionLogons where shopid=" & ishopid
set rsdetail=server.createobject("adodb.recordset")
rsDetail.open strDetail,cnlist

if not rsdetail.eof then
	strshop=rsdetail.fields("name")
end if
rsdetail.close


strDetail = "select * from tblCollisionDetailToyAll inner join tblCollisionDetailToy on tblCollisionDetailToyAll.JobID=tblCollisionDetailToy.JobID where tblCollisionDetailToy.JobID=" & jobid & " "
set rsdetail=server.createobject("adodb.recordset")
rsDetail.open strDetail,cnlist
if rsdetail.eof then
	Response.Write "There is no such record." & "<br>"
	Response.Write "<br>"
else 
	response.write "<b>" & strshop & "</b><br>"	
	'Response.Write rsdetail.Fields("claimq")
%>
<table cellspacing=1 cellpadding=1 width=100% border=0>
<tr valign=top>
<td colspan=2><b><font color=blue><%Response.write rsdetail.Fields("FirstName") & " " & rsdetail.Fields("LastName")%></b></FONT></td>
<td width=17% valign=bottom><font size=2>&nbsp;</font></td>
<td valign=bottom width=5%><font size=2><b></b></font></td>
<td width=40%></td>
</tr>
<tr valign=top>
<td width=15%><font size=2>RO: <%Response.write rsdetail.Fields("ro")%></font></td>
<td width=15%><font size=2><%Response.write "$ " & rsdetail.Fields("amount")%></font></td>
<td><font size=2>1. Quality</font></td>
<td valign=top><font size=2><b><%Response.write rsdetail.Fields("Quality")%></b></font></td>
<td><font size=2><%Response.write "Body Tech:  " & rsdetail.Fields("BodyTech")%></font></td>
</tr>
<tr valign=top>
<td><font size=2>H: <%Response.write rsdetail.Fields("homephone")%></font></td>
<td><font size=2>W: <%Response.write rsdetail.Fields("workphone")%></font></td>
<td><font size=2>2. Treated</font></td>
<td valign=top><font size=2><b><%Response.write rsdetail.Fields("Treated")%></b></font></td>
<td><font size=2><%Response.write "Service Writer: " & rsdetail.Fields("Estimator")%></font></td>
</tr>
<tr valign=top>
<td><font size=2>Year/Make:</font></td>
<td><font size=2><%Response.write rsdetail.Fields("year")%>&nbsp;<%Response.write rsdetail.Fields("make")%></font></td>
<td><font size=2>3. On Time</font></td>
<td valign=top><font size=2><b><%Response.write rsdetail.Fields("OnTime")%></b></font></td>
</tr>
<tr valign=top>
<td><font size=2>Hot Sheet? </font></td>
<td><font size=2 color=red><b><%Response.write rsdetail.Fields("HotSheet")%></b></td>
<td><font size=2>4. Refer Shop</font></td>
<td valign=top><font size=2><b><%Response.write rsdetail.Fields("recshop")%></b></font></td>
</tr>
<tr valign=top>
<td><font size=2>Last Connected Call:</font></td>
<td><font size=2><%Response.write rsdetail.Fields("lastcall")%></font></td>
<td><font size=2><%Response.write rsdetail.Fields("comments")%></font></td>
</tr>
<%Response.Flush%>
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
