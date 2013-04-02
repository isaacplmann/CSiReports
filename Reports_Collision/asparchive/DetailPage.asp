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
		<IMG style="WIDTH: 278px; HEIGHT: 77px" height=44 hspace=1 src="http://www.csicomplete.com/images/logoreport.jpg" vspace=1 border=0>
		</td>
		<td align="center">
		<IMG style="WIDTH: 50px; HEIGHT: 30px" height=44 hspace=1 src="http://www.csicomplete.com/images/email.jpg" vspace=1 border=0>
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
<TABLE cellSpacing=1 cellPadding=1 width=100% border=0>
<TR valign=top>
<TD colspan=2><b><font color=blue><%Response.write rsdetail.Fields("FirstName") & " " & rsdetail.Fields("LastName")%></b></FONT></TD>
<TD width=17% valign=bottom><FONT size=2>1. Greeted</FONT></TD>
<TD valign=bottom Width=5%><FONT size=2><b><%if rsdetail.Fields("greetedq")=-1 then Response.write "YES" else Response.Write "NO"%></b></FONT></TD>
<td width=17%></td>
<TD valign=bottom Width=5%><FONT size=2 color=red><%Response.write rsdetail.Fields("greetedb")%></FONT></TD>
<TD valign=bottom Width=33%><FONT size=2 color=red><%Response.write rsdetail.Fields("greetedc")%></FONT></TD>
</tr>
<TR valign=top>
<TD width=15%><font size=2>RO: <%Response.write rsdetail.Fields("ro")%></FONT></TD>
<TD width=15%><font size=2><%Response.write rsdetail.Fields("referral")%></FONT></TD>
<TD><FONT size=2>2. Quality - Body</FONT></TD>
<TD valign=top><FONT size=2><b><%if rsdetail.Fields("bodyq")=-1 then Response.write "YES" else Response.write "NO"%></b></FONT></TD>
<td><FONT size=2><%Response.write rsdetail.Fields("Body")%></FONT></TD>
<TD valign=top><FONT size=2 color=red><%Response.write rsdetail.Fields("bodyb")%></FONT></TD>
<TD valign=top><FONT size=2 color=red><%Response.write rsdetail.Fields("bodyc")%></FONT></TD>
</tr>
<TR valign=top>
<TD><font size=2>H: <%Response.write rsdetail.Fields("homephone")%></FONT></TD>
<TD><font size=2>W: <%Response.write rsdetail.Fields("workphone")%></FONT></TD>
<TD><FONT size=2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- Paint</FONT></TD>
<TD valign=top><FONT size=2><b><%if rsdetail.Fields("paintq") = -1 then Response.write "YES" else Response.write "NO"%></b></FONT></TD>
<td><FONT size=2><%Response.write rsdetail.Fields("Paint")%></FONT></TD>
<TD valign=top><FONT size=2 color=red><%Response.write rsdetail.Fields("paintb")%></FONT></TD>
<TD valign=top><FONT size=2 color=red><%Response.write rsdetail.Fields("paintc")%></FONT></TD>
</tr>
<TR valign=top>
<TD><font size=2>$ <%Response.write rsdetail.Fields("amount")%></FONT></TD>
<TD><font size=2><%Response.write rsdetail.Fields("insurer")%></FONT></TD>
<TD><FONT size=2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- Mech</FONT></TD>
<TD valign=top><FONT size=2><b><%if rsdetail.Fields("mechq")=-1 then Response.write "YES" else Response.write "NO"%></b></FONT></TD>
<td></td>
<TD valign=top><FONT size=2 color=red><%Response.write rsdetail.Fields("mechb")%></FONT></TD>
<TD valign=top><FONT size=2 color=red><%Response.write rsdetail.Fields("mechc")%></FONT></TD>
</tr>
<TR valign=top>
<TD><font size=2>Year/Make:</FONT></TD>
<TD><font size=2><%Response.write rsdetail.Fields("year")%>&nbsp;<%Response.write rsdetail.Fields("make")%></FONT></TD>
<TD><FONT size=2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;- Detail</FONT></TD>
<TD valign=top><FONT size=2><b><%if rsdetail.Fields("detailq")=-1 then Response.write "YES" else Response.write "NO"%></b></FONT></TD>
<td></td>
<TD valign=top><FONT size=2 color=red><%Response.write rsdetail.Fields("detailb")%></FONT></TD>
<TD valign=top><FONT size=2 color=red><%Response.write rsdetail.Fields("detailc")%></FONT></TD>
</tr>
<TR valign=top>
<TD><font size=2>Hot Sheet? </FONT></TD>
<TD><font size=2 color=red><b><%Response.write rsdetail.Fields("hotsheetdate")%></b></td>
<TD><FONT size=2>3. Cleanliness</FONT></TD>
<TD valign=top><FONT size=2><b><%if rsdetail.Fields("cleanq")=-1 then Response.write "YES" else Response.write "NO"%></b></FONT></TD>
<td></td>
<TD valign=top><FONT size=2 color=red><%Response.write rsdetail.Fields("cleanb")%></FONT></TD>
<TD valign=top><FONT size=2 color=red><%Response.write rsdetail.Fields("cleanc")%></FONT></TD>
</tr>
<TR valign=top>
<TD><font size=2>Last Call Date:</FONT></TD>
<TD><font size=2><%Response.write rsdetail.Fields("lastcall")%></FONT></TD>
<TD><FONT size=2>4. Service</FONT></TD>
<TD valign=top><FONT size=2><b><%if rsdetail.Fields("serviceq")=-1 then Response.write "YES" else Response.write "NO"%></b></FONT></TD>
<td><FONT size=2><%Response.write rsdetail.Fields("estimator")%></td>
<TD valign=top><FONT size=2 color=red><%Response.write rsdetail.Fields("serviceb")%></FONT></TD>
<TD valign=top><FONT size=2 color=red><%Response.write rsdetail.Fields("servicec")%></FONT></TD>
</tr>
<TR valign=top>
<TD colspan=2><FONT size=2>RO Transferred From</FONT></TD>
<TD><FONT size=2>5. Communication</FONT></TD>
<TD valign=top><FONT size=2><b><%if rsdetail.Fields("commq")=-1 then Response.write "YES" else Response.write "NO"%></b></FONT></TD>
<td></FONT></TD>
<TD valign=top><FONT size=2 color=red><%Response.write rsdetail.Fields("commb")%></FONT></TD>
<TD valign=top><FONT size=2 color=red><%Response.write rsdetail.Fields("commc")%></FONT></TD>
</tr>
<TR valign=top>
<TD colspan=2 align=center><font size=2>Previous Month: <%Response.write strRolledOver%></FONT></TD>
<TD><FONT size=2>6. On Time</FONT></TD>
<TD valign=top><FONT size=2><b><%if rsdetail.Fields("ontimeq")=-1 then Response.write "YES" else Response.write "NO"%></b></FONT></TD>
<td></FONT></TD>
<TD valign=top><FONT size=2 color=red><%Response.write rsdetail.Fields("ontimeb")%></FONT></TD>
<TD valign=top><FONT size=2 color=red><%Response.write rsdetail.Fields("ontimec")%></FONT></TD>
</tr>
<TR valign=top>
<TD rowspan=3 colspan=2 valign=top><FONT size=2 color=blue><%Response.write rsdetail.Fields("testimony")%></FONT></TD>
<TD><FONT size=2>7. 2nd Return Visit</FONT></TD>
<TD valign=top><FONT size=2><b><%if rsdetail.Fields("returnq")=-1 then Response.write "YES" else Response.write "NO"%></b></FONT></TD>
<td></FONT></TD>
<TD valign=top><FONT size=2 color=red><%Response.write rsdetail.Fields("returnb")%></FONT></TD>
<TD valign=top><FONT size=2 color=red><%Response.write rsdetail.Fields("returnc")%></FONT></TD>
</tr>
<TR valign=top>
<TD><FONT size=2>8. Refer Shop</FONT></TD>
<TD valign=top><FONT size=2><b><%if rsdetail.Fields("recshopq")=-1 then Response.write "YES" else Response.write "NO"%></b></FONT></TD>
<td></FONT></TD>
<TD valign=top><FONT size=2 color=red></FONT></TD>
<TD valign=top><FONT size=2 color=red><%Response.write rsdetail.Fields("recshopc")%></FONT></TD>
</tr>
<TR valign=top>
<TD><FONT size=2>9. Handle Claim</FONT></TD>
<TD valign=top><FONT size=2><b><%Response.write strclaimq%></b></FONT></TD>
<td><%
%></FONT></TD>
<TD valign=top><FONT size=2></FONT></TD>
<TD valign=top><FONT size=2 color=red><%Response.write rsdetail.Fields("claimc")%></FONT></TD>
</tr>
<TR valign=top>
<TD></FONT></TD>
<TD></FONT></TD>
<TD><FONT size=2>10. Rec Ins.</FONT></TD>
<TD valign=top><FONT size=2><b><%Response.write strrecinsq%></b></FONT></TD>
<td></FONT></TD>
<TD valign=top><FONT size=2></FONT></TD>
<TD valign=top><FONT size=2 color=red><%Response.write rsdetail.Fields("recinsc")%></FONT></TD>
</tr>
<TR valign=top>
<TD></FONT></TD>
<TD></FONT></TD>
<TD><FONT size=2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Claimant</FONT></TD>
<TD valign=top><FONT size=2><b><%if rsdetail.Fields("claimant")=-1 then Response.write "YES" else Response.write "NO"%></b></FONT></TD>
<td></FONT></TD>
<td></FONT></TD>
<td></FONT></TD>
</tr>
<TR valign=top>
<TD></FONT></TD>
<TD></FONT></TD>
<TD colspan=5><font size=2 color=red><%Response.write rsdetail.Fields("comments")%></FONT></TD>
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
