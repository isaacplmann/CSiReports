<%@ Language=VBScript %>
<%Response.Buffer=True%>
<html>
<head>
<title>RO List</title>
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
<%
dim iShopID 
dim dRptMonth
dim strDetail
dim iSort
dim strMonth
dim strYear
dim strSQL
dim strShop

if session("isort")=1 then 
	session("isort")=2
else
	session("isort")=1
end if

ishopid=Request.QueryString("shopid")
dRptMonth=Request.QueryString("Month")
strSort=Request.QueryString("Sort")

stryear = year(drptmonth)
stryear = right(stryear,2)
strmonth=MonthName(DatePart("m", drptmonth),True)
strmonth = strmonth & " " & stryear
response.write "<TABLE width=""100%"">"

set cnList=server.createobject("adodb.Connection")
cnList.open session("Constring")

strdetail = "Select name from tblCollisionLogons where shopid=" & ishopid
set rsdetail=server.createobject("adodb.recordset")
rsDetail.open strDetail,cnlist

if not rsdetail.eof then
	strshop=rsdetail.fields("name")
end if
rsdetail.close
if session("clienttype")<>"S" then
	Response.Write "<TR>"
	Response.Write "<Td align=left><font size=3 color=brown><strong>"
	Response.Write strshop
	Response.Write "</strong></font></Td>"
	Response.Write "</TR>"
end if
Response.Write "<TR>"
Response.Write "<Td align=left><font size=3 color=brown><strong>"
Response.Write strmonth
Response.Write "</strong></font></Td>"
Response.Write "</TR>"
Response.Write "</table>"


strDetail = "select tblCollisionDetail.* from tblCollisionDetail where tblCollisionDetail.ShopID=" & ishopid & " and RptMonth='" & dRptMonth & "' order by "

strsort = ucase(strsort)
select case strSort
	case "STATUS"
		if session("isort")=1 then
			strDetail = strdetail & "statusid, RecShopNum desc, lastname"
		else
			strDetail = strdetail & "statusid desc, RecShopNum desc, lastname"		
		end if
	case "RO"
		if session("isort")=1 then
			strDetail = strdetail & "RO"
		else
			strDetail = strdetail & "RO desc"
		end if
	case "LAST"
		if session("isort")=1 then
			strDetail = strdetail & "LastName"
		else
			strDetail = strdetail & "LastName desc"		
		end if
	case "FIRST"
		if session("isort")=1 then
			strDetail = strdetail & "FirstName"
		else
			strDetail = strdetail & "FirstName desc"
		end if
	case "SW"
		if session("isort")=1 then
			strDetail = strdetail & "Estimator"
		else
			strDetail = strdetail & "Estimator desc"
		end if
	case "LASTCALL"
		if session("isort")=1 then
			strDetail = strdetail & "LastCall desc"
		else
			strDetail = strdetail & "LastCall"
		end if
	case "RECSHOPNUM"
		if session("isort")=1 then
			strDetail = strdetail & "RecShopNum"
		else
			strDetail = strdetail & "RecShopNum desc"
		end if
	case "HS"
		if session("isort")=1 then
			strDetail = strdetail & "HotSheetDate desc"
		else
			strDetail = strdetail & "HotSheetDate"		
		end if
	case "INS"
		if session("isort")=1 then
			strDetail = strdetail & "Insurer"
		else
			strDetail = strdetail & "Insurer desc"		
		end if
	case else
		session("isort")=1
		strDetail = strdetail & "statusid, RecShopNum desc, lastname"
end select

	set rsdetail=server.createobject("adodb.recordset")
	rsDetail.open strDetail,cnlist

	if rsdetail.eof then
	response.write "There are no ROs to display for the month."
	response.write "<br>"
	else
	Response.Write "<P align = center><Font Size=2> "
	Response.Write "<IMG Style = ""WIDTH: 15px; HEIGHT: 15px"" height=44 hspace=1 src=""http://www.csicomplete.com/images/green.gif"" vspace=1 border=0>"
	Response.Write "Completed Survey&nbsp;&nbsp;&nbsp;&nbsp;"
	Response.Write "<IMG Style = ""WIDTH: 15px; HEIGHT: 15px"" height=44 hspace=1 src=""http://www.csicomplete.com/images/yellow.gif"" vspace=1 border=0>"
	Response.Write "Pending Survey&nbsp;&nbsp;&nbsp;&nbsp;"
	Response.Write "<IMG Style = ""WIDTH: 15px; HEIGHT: 15px"" height=44 hspace=1 src=""http://www.csicomplete.com/images/red.gif"" vspace=1 border=0>"
	Response.Write "Unreachable Survey "
	Response.Write "</font></p>"
	
	response.write "<TABLE width=""100%"">"
	response.write "<TH align=left><font size=2><a href='ROList_gerber.asp?ShopID=" & ishopid & "&Month=" & drptmonth & "&Sort=Status'>Status</a></TH>"
	response.write "<TH align=center><font size=2><a href='ROList_gerber.asp?ShopID=" & ishopid & "&Month=" & drptmonth & "&Sort=RO'>RO</a></TH>"
	response.write "<TH align=left><font size=2>View</TH>"
	response.write "<TH align=left><font size=2><a href='ROList_gerber.asp?ShopID=" & ishopid & "&Month=" & drptmonth & "&Sort=Last'>Last</a></TH>"
	response.write "<TH align=left><font size=2><a href='ROList_gerber.asp?ShopID=" & ishopid & "&Month=" & drptmonth & "&Sort=First'>First</a></TH>"
	response.write "<TH align=left><font size=2><a href='ROList_gerber.asp?ShopID=" & ishopid & "&Month=" & drptmonth & "&Sort=RecShopNum'>0 - 10</a></TH>"
	response.write "<TH align=left><font size=2><a href='ROList_gerber.asp?ShopID=" & ishopid & "&Month=" & drptmonth & "&Sort=HS'>Hot Sheet</a></TH>"
	response.write "<TH align=left><font size=2><a href='ROList_gerber.asp?ShopID=" & ishopid & "&Month=" & drptmonth & "&Sort=SW'>Estimator</a></TH>"
	response.write "<TH align=left><font size=2><a href='ROList_gerber.asp?ShopID=" & ishopid & "&Month=" & drptmonth & "&Sort=LastCall'>Last Connected Call</a></TH>"
	response.write "<TH align=left><font size=2><a href='ROList_gerber.asp?ShopID=" & ishopid & "&Month=" & drptmonth & "&Sort=Ins'>Insurer</a></TH>"
	response.write "</tr>"


	do until rsdetail.eof


	dim ParentID
	dim Name
	ParentID = session("sessionparentid")
	MSO = session("sessionmso")
	Name = session("sessionname")
	dim pic
	
	'pic="<IMG Style = ""WIDTH: 15px; HEIGHT: 15px"" height=44 hspace=1 src=""http://www.csicomplete.com/images/green.gif"" vspace=1 border=0>"
	'strURL = "http://www.csicomplete.com/reports_collision/detailpagemso_gerber.asp?jobid=" & rsdetail.fields("jobid") & "&ShopID=" & ishopid & " "
	
	select case rsdetail.fields("statusid")
	    case 1
	    pic="<IMG Style = ""WIDTH: 15px; HEIGHT: 15px"" height=44 hspace=1 src=""http://www.csicomplete.com/images/green.gif"" vspace=1 border=0>"
	    strURL = "http://www.csicomplete.com/reports_collision/detailpagemso_gerber.asp?jobid=" & rsdetail.fields("jobid") & "&ShopID=" & ishopid & " "
	    case 2
	    pic="<IMG Style = ""WIDTH: 15px; HEIGHT: 15px"" height=44 hspace=1 src=""http://www.csicomplete.com/images/yellow.gif"" vspace=1 border=0>"
	    strURL = "http://www.csicomplete.com/reports_collision/DetailPageInc.asp?jobid=" & rsdetail.fields("jobid") & "&ShopID=" & ishopid & " "
	    Case else
	    pic="<IMG Style = ""WIDTH: 15px; HEIGHT: 15px"" height=44 hspace=1 src=""http://www.csicomplete.com/images/red.gif"" vspace=1 border=0>"
	    strURL = "http://www.csicomplete.com/reports_collision/DetailPageInc.asp?jobid=" & rsdetail.fields("jobid") & "&ShopID=" & ishopid & " "
	    end select
	
	Response.write "<tr>"
	response.write "<td align=center><font size=2>" & pic & "</td>"
	response.write "<td><font size=2>" & rsdetail.fields("RO") & "</td>"
	response.write "<td><font size=2>"
	Response.Write "<a href = ""javascript: openZoom1('" & strURL & "','SURVEY');"">Details...</a>"
	Response.write "</td>"
	response.write "<td><font size=2>" & rsdetail.fields("LastName") & "</td>"
	response.write "<td><font size=2>" & rsdetail.fields("FirstName") & "</td>"
	
	
	
	If IsNumeric(rsdetail.fields("RecShopNum")) then 
		response.write "<td><font size=2>" & round(rsdetail.fields("RecShopNum"),1) & "</td>"
	Else
		response.write "<td><font size=2></td>"
	End If
	
	
	response.write "<td><font color=red size=2>" & rsdetail.fields("HotSheetDate") & "</td>"
	response.write "<td><font size=2>" & rsdetail.fields("Estimator") & "</td>"
	response.write "<td><font size=2>" & rsdetail.fields("LastCall") & "</td>"
	response.write "<td><font size=2>" & rsdetail.fields("Insurer") & "</td>"
	Response.write "</tr>"
	rsdetail.movenext
	loop

	response.write "</TABLE>"
	response.write "<P>"
	Response.Write "<P align = center><Font Size=2> "
	Response.Write "<IMG Style = ""WIDTH: 15px; HEIGHT: 15px"" height=44 hspace=1 src=""http://www.csicomplete.com/images/green.gif"" vspace=1 border=0>"
	Response.Write "Completed Survey&nbsp;&nbsp;&nbsp;&nbsp;"
	Response.Write "<IMG Style = ""WIDTH: 15px; HEIGHT: 15px"" height=44 hspace=1 src=""http://www.csicomplete.com/images/yellow.gif"" vspace=1 border=0>"
	Response.Write "Pending Survey&nbsp;&nbsp;&nbsp;&nbsp;"
	Response.Write "<IMG Style = ""WIDTH: 15px; HEIGHT: 15px"" height=44 hspace=1 src=""http://www.csicomplete.com/images/red.gif"" vspace=1 border=0>"
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
