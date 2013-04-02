<%@ Language=VBScript %>
<%Response.Buffer=True%>
<HTML>
<HEAD>
<TITLE>RO List</TITLE>
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
<%
dim iShopID 
dim dRptMonth

ishopid=Request.QueryString("shopid")
dRptMonth=Request.QueryString("Month")
set cnList=server.createobject("adodb.Connection")
cnList.open session("Constring")
strDetail = "select * from tblCollisionDetail where ShopID=" & ishopid & " and RptMonth='" & dRptMonth & "' order by statusid,CSi desc, lastname"
set rsdetail=server.createobject("adodb.recordset")
rsDetail.open strDetail,cnlist

if rsdetail.eof then
response.write "There are no ROs to display for the month."
response.write "<br>"
else
Response.Write "<P align = center><Font Size=2> "
Response.Write "<IMG Style = ""WIDTH: 15px; HEIGHT: 15px"" height=44 hspace=1 src=""http://209.115.48.235/images/green.gif"" vspace=1 border=0>"
Response.Write "Completed Survey&nbsp;&nbsp;&nbsp;&nbsp;"
Response.Write "<IMG Style = ""WIDTH: 15px; HEIGHT: 15px"" height=44 hspace=1 src=""http://209.115.48.235/images/yellow.gif"" vspace=1 border=0>"
Response.Write "Pending Survey&nbsp;&nbsp;&nbsp;&nbsp;"
Response.Write "<IMG Style = ""WIDTH: 15px; HEIGHT: 15px"" height=44 hspace=1 src=""http://209.115.48.235/images/red.gif"" vspace=1 border=0>"
Response.Write "Unreachable Survey "
Response.Write "</font></p>"
response.write "<TABLE width=""100%"">"
response.write "<TR>"
response.write "<TH align=left>Status</TH>"
response.write "<TH align=center>RO</TH>"
response.write "<TH align=left>View Survey</TH>"
response.write "<TH align=left>Last Name</TH>"
response.write "<TH align=left>First Name</TH>"
response.write "<TH align=left>CSi</TH>"
response.write "<TH align=left>Hot Sheet</TH>"
response.write "<TH align=left>Estimator</TH>"
response.write "<TH align=left>Last Call</TH>"
response.write "</tr>"
do until rsdetail.eof

dim csi
CSI = rsdetail.Fields("CSI")*100
if csi>=0 then csi=csi & "%"
dim pic
select case rsdetail.fields("statusid")
case 1
pic="<IMG Style = ""WIDTH: 15px; HEIGHT: 15px"" height=44 hspace=1 src=""http://209.115.48.235/images/green.gif"" vspace=1 border=0>"
strURL = "http://www.csicomplete.com/reports_collision/DetailPage.asp?jobid=" & rsdetail.fields("jobid") & "&ShopName=" & session("sessionname") & " "
case 2
pic="<IMG Style = ""WIDTH: 15px; HEIGHT: 15px"" height=44 hspace=1 src=""http://209.115.48.235/images/yellow.gif"" vspace=1 border=0>"
strURL = "http://www.csicomplete.com/reports_collision/DetailPageInc.asp?jobid=" & rsdetail.fields("jobid") & "&ShopName=" & session("sessionname") & " "
Case else
pic="<IMG Style = ""WIDTH: 15px; HEIGHT: 15px"" height=44 hspace=1 src=""http://209.115.48.235/images/red.gif"" vspace=1 border=0>"
strURL = "http://www.csicomplete.com/reports_collision/DetailPageInc.asp?jobid=" & rsdetail.fields("jobid") & "&ShopName=" & session("sessionname") & " "
end select
Response.write "<tr>"
response.write "<td align=center><font size=2>" & pic & "</td>"
response.write "<td><font size=2>" & rsdetail.fields("RO") & "</td>"
response.write "<td><font size=2>"
Response.Write "<a href = ""javascript: openZoom1('" & strURL & "','SURVEY');"">Details...</a>"
Response.write "</td>"
response.write "<td><font size=2>" & rsdetail.fields("LastName") & "</td>"
response.write "<td><font size=2>" & rsdetail.fields("FirstName") & "</td>"
response.write "<td><font size=2>" & csi & "</td>"
response.write "<td><font color=red size=2>" & rsdetail.fields("HotSheetDate") & "</td>"
response.write "<td><font size=2>" & rsdetail.fields("Estimator") & "</td>"
response.write "<td><font size=2>" & rsdetail.fields("LastCall") & "</td>"
Response.write "</tr>"
rsdetail.movenext
loop
response.write "</TABLE>"
response.write "<P>"
Response.Write "<P align = center><Font Size=2> "
Response.Write "<IMG Style = ""WIDTH: 15px; HEIGHT: 15px"" height=44 hspace=1 src=""http://209.115.48.235/images/green.gif"" vspace=1 border=0>"
Response.Write "Completed Survey&nbsp;&nbsp;&nbsp;&nbsp;"
Response.Write "<IMG Style = ""WIDTH: 15px; HEIGHT: 15px"" height=44 hspace=1 src=""http://209.115.48.235/images/yellow.gif"" vspace=1 border=0>"
Response.Write "Pending Survey&nbsp;&nbsp;&nbsp;&nbsp;"
Response.Write "<IMG Style = ""WIDTH: 15px; HEIGHT: 15px"" height=44 hspace=1 src=""http://209.115.48.235/images/red.gif"" vspace=1 border=0>"
Response.Write "Unreachable Survey "
Response.Write "</font></p>"
end if
response.write "<TABLE cellSpacing=1 cellPadding=1 width=100% border=0>"
response.write "</TABLE>"
response.write "<P>"
response.write "<CENTER><STRONG><EM><font size=1>"
response.write "Information compiled by CSi Complete, a national provider of CSI and call center services to business. For more information please call (800) 343-0641. Copyright 2003 CSi Complete</FONT></B></I></TD>"
response.write "</font></EM></STRONG></CENTER>"
response.write "</BODY></HTML>"
%>
