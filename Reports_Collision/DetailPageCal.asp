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
<font face=Arial>
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

</font><br>
<%
Response.Flush
dim ConnString
dim jobid
dim shopname
dim strRentalQ

ConnString= session("ConString")
jobid = Request.QueryString("jobid")
'shopname = Request.QueryString("shopname")
iShopID = Request.QueryString("shopid")

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

	if rsdetail.Fields("rentalq") = -1 then
		strRentalQ="YES"
	elseif rsdetail.Fields("rentalq")=0 then
		strRentalQ="NO"
	else
		strRentalQ="N/A"
	end if

	if rsdetail.Fields("rolledover") = -1 then
		strRolledOver="YES"
	elseif rsdetail.Fields("rolledover")=0 then
		strRolledOver="NO"
	else
		strRolledOver="NO"
	end if

%>
<table cellspacing=1 cellpadding=1 width=100% border=0>
<tr valign=top>
<td colspan=2><b><font color=blue><%Response.write rsdetail.Fields("FirstName") & " " & rsdetail.Fields("LastName")%></b></FONT></td>
<td width=22% valign=bottom><font size=2>1. Quality - Body</font></td>
<td valign=bottom width=5%><font size=2><b><%if rsdetail.Fields("bodyq")=-1 then Response.write "YES" else Response.write "NO"%></b></font></td>
<td valign=bottom width=17%><font size=2><%Response.write rsdetail.Fields("Body")%></font></td>
<td valign=bottom width=5%><font size=2 color=red><%Response.write rsdetail.Fields("bodyb")%></font></td>
<td valign=bottom width=33%><font size=2 color=red><%Response.write rsdetail.Fields("bodyc")%></font></td>
</tr>
<tr valign=top>
<td width=15%><font size=2>RO: <%Response.write rsdetail.Fields("ro")%></font></td>
<td width=15%></td>
<td><font size=2>&nbsp;&nbsp;&nbsp;&nbsp;Quality - Paint</font></td>
<td valign=top><font size=2><b><%if rsdetail.Fields("paintq")=-1 then Response.write "YES" else Response.write "NO"%></b></font></td>
<td><font size=2><%Response.write rsdetail.Fields("Paint")%></font></td>
<td valign=top><font size=2 color=red><%Response.write rsdetail.Fields("paintb")%></font></td>
<td valign=top><font size=2 color=red><%Response.write rsdetail.Fields("paintc")%></font></td>
</tr>
<tr valign=top>
<td><font size=2>H: <%Response.write rsdetail.Fields("homephone")%></font></td>
<td><font size=2>W: <%Response.write rsdetail.Fields("workphone")%></font></td>
<td><font size=2>2. Improve Quality</font></td>
<td></td>
<td></td>
<td valign=top></td>
<td valign=top><font size=2 color=red><%Response.write rsdetail.Fields("improvequal")%></font></td>
</tr>
<tr valign=top>
<td><font size=2>$ <%Response.write rsdetail.Fields("amount")%></font></td>
<td><font size=2><%Response.write rsdetail.Fields("insurer")%></font></td>
<td><font size=2>3. Service</font></td>
<td valign=top><font size=2><b><%if rsdetail.Fields("serviceq")=-1 then Response.write "YES" else Response.write "NO"%></b></font></td>
<td><font size=2><%Response.write rsdetail.Fields("estimator")%></font></td>
<td valign=top><font size=2 color=red><%Response.write rsdetail.Fields("serviceb")%></font></td>
<td valign=top><font size=2 color=red><%Response.write rsdetail.Fields("servicec")%></font></td>
</tr>
<tr valign=top>
<td><font size=2>Year/Make:</font></td>
<td><font size=2><%Response.write rsdetail.Fields("year")%>&nbsp;<%Response.write rsdetail.Fields("make")%></font></td>
<td><font size=2>4. Improve Service</font></td>
<td></td>
<td></td>
<td></td>
<td valign=top><font size=2 color=red><%Response.write rsdetail.Fields("improveserv")%></font></td>
</tr>
<tr valign=top>
<td><font size=2>Hot Sheet? </font></td>
<td><font size=2 color=red><b><%Response.write rsdetail.Fields("hotsheetdate")%></b></td>
<td><font size=2>5. Informed</font></td>
<td valign=top><font size=2><b><%if rsdetail.Fields("commq")=-1 then Response.write "YES" else Response.write "NO"%></b></font></td>
<td></td>
<td valign=top><font size=2 color=red><%Response.write rsdetail.Fields("commb")%></font></td>
<td valign=top><font size=2 color=red><%Response.write rsdetail.Fields("commc")%></font></td>
</tr>
<tr valign=top>
<td><font size=2>Last Connected Call:</font></td>
<td><font size=2><%Response.write rsdetail.Fields("lastcall")%></font></td>
<td><font size=2>6. On Time</font></td>
<td valign=top><font size=2><b><%if rsdetail.Fields("ontimeq")=-1 then Response.write "YES" else Response.write "NO"%></b></font></td>
<td></td>
<td valign=top><font size=2 color=red><%Response.write rsdetail.Fields("ontimeb")%></font></td>
<td valign=top><font size=2 color=red><%Response.write rsdetail.Fields("ontimec")%></font></td>
</tr>
<tr valign=top>
<td colspan=2><font size=2>RO Transferred From</font></td>
<td><font size=2>7. Recommend Shop</font></td>
<td valign=top><font size=2><b><%if rsdetail.Fields("recshopq")=-1 then Response.write "YES" else Response.write "NO"%></b></font></td>
<td></td>
<td></td>
<td valign=top><font size=2 color=red><%Response.write rsdetail.Fields("recshopc")%></font></td>
</tr>
<tr valign=top>
<td colspan=2 align=center><font size=2>Previous Month: <%Response.write strrolledover%></font></td>
<td><font size=2>8. Recommend Ins.</font></td>
<td valign=top><font size=2><b><%if rsdetail.Fields("recinsq")=-1 then Response.write "YES" else Response.write "NO"%></b></font></td>
<td></td>
<td></td>
<td valign=top><font size=2 color=red><%Response.write rsdetail.Fields("recinsc")%></font></td>
</tr>
<tr valign=top>
<td rowspan=3 colspan=2 valign=top><font size=2 color=blue><%Response.write rsdetail.Fields("testimony")%></font></td>
<td><font size=2>9. Rental Car</font></td>
<td valign=top><font size=2><b><%Response.write strrentalq%></b></font></td>
<td></td>
<td></td>
<td></td>
</tr>
<tr valign=top>
<td><font size=2>&nbsp;&nbsp;&nbsp;Claimant</font></td>
<td valign=top><font size=2><b><%if rsdetail.Fields("claimant")=-1 then Response.write "YES" else Response.write "NO"%></b></font></td>
<td></FONT></td>
<td></FONT></td>
<td></FONT></td>
</tr>
<tr valign=top>
<td></FONT></td>
<td></FONT></td>
<td colspan=5><font size=2 color=red><%Response.write rsdetail.Fields("comments")%></font></td>
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
