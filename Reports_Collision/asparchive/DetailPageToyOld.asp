<%@ Language=VBScript %>
<%Response.Buffer=True%>
<HTML>
<HEAD>
<TITLE>Survey Detail</TITLE>
<SCRIPT LANGUAGE="Javascript">
function openZoom1(url, name) {
  popupWin = window.open(url, name,
"ScrollBars = yes, resizable=no,location=no,toolbar=no,top=6,left=155,width=750,height=470");
}
<!--
if (window.focus) {
  self.focus();
}
//-->
</SCRIPT>
</HEAD>
<BODY><font face=arial>
<FONT face=Arial><A HREF="http://www.csicomplete.com">
<table width=100% border=0>
	<tr>
		<td>
		<IMG style="WIDTH: 278px; HEIGHT: 77px" height=44 hspace=1 src="http://209.115.48.235/images/logoreport.jpg" vspace=1 border=0>
		</td>
		<td align="center">
		<IMG style="WIDTH: 50px; HEIGHT: 30px" height=44 hspace=1 src="http://209.115.48.235/images/email.jpg" vspace=1 border=0>
		</td>
	</tr>
</table>

</A></FONT><br>
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
<TABLE cellSpacing=1 cellPadding=1 width=100% border=0>
<TR valign=top>
<TD colspan=2><b><font color=blue><%Response.write rsdetail.Fields("FirstName") & " " & rsdetail.Fields("LastName")%></b></FONT></TD>
<TD width=17% valign=bottom><FONT size=2></FONT></TD>
<TD valign=bottom Width=5%><FONT size=2><b></b></FONT></TD>
<td width=40%></td>
</tr>
<TR valign=top>
<TD width=15%><font size=2>RO: <%Response.write rsdetail.Fields("ro")%></FONT></TD>
<TD width=15%><font size=2><%Response.write "$ " & rsdetail.Fields("amount")%></FONT></TD>
<TD><FONT size=2>1. Quality</FONT></TD>
<TD valign=top><FONT size=2><b><%Response.write rsdetail.Fields("Quality")%></b></FONT></TD>
<td><FONT size=2><%Response.write "Body Tech:  " & rsdetail.Fields("BodyTech")%></FONT></TD>
</tr>
<TR valign=top>
<TD><font size=2>H: <%Response.write rsdetail.Fields("homephone")%></FONT></TD>
<TD><font size=2>W: <%Response.write rsdetail.Fields("workphone")%></FONT></TD>
<TD><FONT size=2>2. Treated</FONT></TD>
<TD valign=top><FONT size=2><b><%Response.write rsdetail.Fields("Treated")%></b></FONT></TD>
<td><FONT size=2><%Response.write "Service Writer: " & rsdetail.Fields("Estimator")%></FONT></TD>
</tr>
<TR valign=top>
<TD><font size=2>Year/Make:</FONT></TD>
<TD><font size=2><%Response.write rsdetail.Fields("year")%>&nbsp;<%Response.write rsdetail.Fields("make")%></FONT></TD>
<TD><FONT size=2>3. On Time</FONT></TD>
<TD valign=top><FONT size=2><b><%Response.write rsdetail.Fields("OnTime")%></b></FONT></TD>
</tr>
<TR valign=top>
<TD><font size=2>Hot Sheet? </FONT></TD>
<TD><font size=2 color=red><b><%Response.write rsdetail.Fields("HotSheet")%></b></td>
<TD><FONT size=2>4. Refer Shop</FONT></TD>
<TD valign=top><FONT size=2><b><%Response.write rsdetail.Fields("recshop")%></b></FONT></TD>
</tr>
<TR valign=top>
<TD><font size=2>Last Call Date:</FONT></TD>
<TD><font size=2><%Response.write rsdetail.Fields("lastcall")%></FONT></TD>
<TD><FONT size=2><%Response.write rsdetail.Fields("comments")%></FONT></TD>
</tr>
<%Response.Flush%>
</TABLE>
<%
end if
%>

<P>
<CENTER><STRONG><EM><font size=1>
Information compiled by CSi Complete, a national provider of CSI and call center services to business.</br>
For more information please call (800) 343-0641. Copyright 2003 CSi Complete</FONT></B></I>
</font></EM></STRONG></CENTER>
</BODY></HTML>
