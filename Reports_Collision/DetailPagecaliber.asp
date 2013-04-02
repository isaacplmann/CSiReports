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
function printPage() {
  if (window.print)
    window.print()
  else
    alert("Sorry, your browser doesn't support this feature.");
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
		<img style="WIDTH: 278px; HEIGHT: 77px" height=44 hspace=1 src="http://www.csicomplete.com/images/logoreport.jpg" vspace=1 border=0>
		</td>
		<td align="center">
		<form>
        <input type="button" value="Print" onClick="printPage()">
        </form>
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

ConnString= session("ConString")
jobid = Request.QueryString("jobid")
iShopID = Request.QueryString("shopid")
'shopname = Request.QueryString("shopname")

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
	'Response.Write rsdetail.Fields("claimq")
	if rsdetail.Fields("claimq") = -1 then
		strClaimQ="YES"
	elseif rsdetail.Fields("claimq")=0 then
		strClaimQ="NO"
	else
		strClaimQ="N/A"
	end if

	if rsdetail.Fields("recinsq") = -1 then
		strRecInsQ="YES"
	elseif rsdetail.Fields("recinsq")=0 then
		strRecInsQ="NO"
	else
		strRecInsQ="N/A"
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
<td colspan=2><b><font color=blue><%Response.write rsdetail.Fields("FirstName") & " " & rsdetail.Fields("LastName")%></font></b></td>
<td width=17% valign=bottom><font size=2>1. Shop Rating</font></td>
<td valign=bottom width=5%><font size=2><b><%Response.write rsdetail.Fields("recshopnum")%></b></font></td>
<td width=17%></td>
<td valign=bottom width=5%>&nbsp;</td>
<td valign=bottom width=33%>&nbsp;</td>
</tr>
<tr valign=top>
<td width=15%><font size=2>RO: <%Response.write rsdetail.Fields("ro")%></font></td>
<td width=15%><font size=2><%Response.write rsdetail.Fields("referral")%></font></td>
<td><font size=2>2. Quality - Body</font></td>
<td valign=top><font size=2><b><%if rsdetail.Fields("bodyq")=-1 then Response.write "YES" else Response.write "NO"%></b></font></td>
<td><font size=2><%Response.write rsdetail.Fields("Body")%></font></td>
<td valign=top><font size=2 color=red><%Response.write rsdetail.Fields("bodyb")%></font></td>
<td valign=top><font size=2 color=red><%Response.write rsdetail.Fields("bodyc")%></font></td>
</tr>
<tr valign=top>
<td><font size=2>H: <%Response.write rsdetail.Fields("homephone")%></font></td>
<td><font size=2>W: <%Response.write rsdetail.Fields("workphone")%></font></td>
<td><font size=2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- Paint</font></td>
<td valign=top><font size=2><b><%if rsdetail.Fields("paintq") = -1 then Response.write "YES" else Response.write "NO"%></b></font></td>
<td><font size=2><%Response.write rsdetail.Fields("Paint")%></font></td>
<td valign=top><font size=2 color=red><%Response.write rsdetail.Fields("paintb")%></font></td>
<td valign=top><font size=2 color=red><%Response.write rsdetail.Fields("paintc")%></font></td>
</tr>
<tr valign=top>
<td><font size=2>$ <%Response.write rsdetail.Fields("amount")%></font></td>
<td><font size=2><%Response.write rsdetail.Fields("insurer")%></font></td>
<td><font size=2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- Mech</font></td>
<td valign=top><font size=2><b><%if rsdetail.Fields("mechq")=-1 then Response.write "YES" else Response.write "NO"%></b></font></td>
<td></td>
<td valign=top><font size=2 color=red><%Response.write rsdetail.Fields("mechb")%></font></td>
<td valign=top><font size=2 color=red><%Response.write rsdetail.Fields("mechc")%></font></td>
</tr>
<tr valign=top>
<td><font size=2>Year/Make:</font></td>
<td><font size=2><%Response.write rsdetail.Fields("year")%>&nbsp;<%Response.write rsdetail.Fields("make")%></font></td>
<td><font size=2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- Detail</font></td>
<td valign=top><font size=2><b><%if rsdetail.Fields("detailq")=-1 then Response.write "YES" else Response.write "NO"%></b></font></td>
<td></td>
<td valign=top><font size=2 color=red><%Response.write rsdetail.Fields("detailb")%></font></td>
<td valign=top><font size=2 color=red><%Response.write rsdetail.Fields("detailc")%></font></td>
</tr>
<tr valign=top>
<td><font size=2>Hot Sheet? </font></td>
<td><font size=2 color=red><b><%Response.write rsdetail.Fields("hotsheetdate")%></b></font></td>
<td><font size=2>3. Informed</font></td>
<td valign=top><font size=2><b><%if rsdetail.Fields("commq")=-1 then Response.write "YES" else Response.write "NO"%></b></font></td>
<td></td>
<td valign=top><font size=2 color=red><%Response.write rsdetail.Fields("commb")%></font></td>
<td valign=top><font size=2 color=red><%Response.write rsdetail.Fields("commc")%></font></td>
</tr>
<tr valign=top>
<td><font size=2>Last Connected Call:</font></td>
<td><font size=2><%Response.write rsdetail.Fields("lastcall")%></font></td>
<td><font size=2>4. On Time</font></td>
<td valign=top><font size=2><b><%if rsdetail.Fields("ontimeq")=-1 then Response.write "YES" else Response.write "NO"%></b></font></td>
    <td>&nbsp;</td>
<td valign=top><font size=2 color=red><%Response.write rsdetail.Fields("ontimeb")%></font></td>
<td valign=top><font size=2 color=red><%Response.write rsdetail.Fields("ontimec")%></font></td>
</tr>
<tr valign=top>
<td colspan=2><font size=2>RO Transferred From</font></td>
<td><font size=2>5. <font size=2 face="arial">&nbsp;2nd Return Visit</font></font></td>
<td valign=top><font size=2><b><%if rsdetail.Fields("returnq")=-1 then Response.write "YES" else Response.write "NO"%></b></font></td>
<td></td>
<td valign=top><font size=2 color=red><%Response.write rsdetail.Fields("returnb")%></font></td>
<td valign=top><font size=2 color=red><%Response.write rsdetail.Fields("returnc")%></font></td>
</tr>
<tr valign=top>
<td colspan=2 align=center><font size=2>Previous Month: <%Response.write strRolledOver%></font></td>
<td><font size=2>6. <font size=2 face="arial">Refer Shop</font></font></td>
<td valign=top><font size=2><b><%if rsdetail.Fields("recshopq")=-1 then Response.write "YES" else Response.write "NO"%></b></font></td>
<td></td>
<td valign=top>&nbsp;</td>
<td valign=top><font size=2 color=red><%Response.write rsdetail.Fields("recshopc")%></font></td>
</tr>

<td colspan=5><font size=2 color=red><%Response.write rsdetail.Fields("comments")%></font></td>
</tr>
<%Response.Flush%>
</table>
<%
end if
%>

<p>
<center>
    <strong><em><font size=1> Information compiled by CSi Complete, a national 
    provider of CSI and call center services to business.</br> For more information 
    please call (800) 343-0641. Copyright 2008 CSi Complete</font></B></I> </em></strong>
</center>
</body></html>
