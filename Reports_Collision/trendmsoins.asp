<%@ Language=VBScript %>
<%Response.Buffer=True%>
<html>
<head>
<title>Trend Report</title>
</head>
<body><font face=arial size=2>
<%
dim iShopID 
dim dRptMonth
dim strMonth
dim strYear
dim MaxMonth
ishopid=Request.QueryString("shopid")
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
	Response.Write "<br>"
	Response.Write "<br>"
end if
set rsdetail=server.createobject("adodb.recordset")
'Get the latest trend month
strDetail = "select max(trendmonth) as MaxMonth from tblCollisionTrend where shopID=" & ishopid
rsDetail.open strDetail, cnlist
if rsdetail.eof then 
	MaxMonth = date()
else
	MaxMonth = rsdetail.fields("MaxMonth")
end if
rsDetail.close


strDetail = "select * from tblCollisionTrend where ShopID=" & ishopid & " and TrendMonth between '" & dateadd("M",-11,MaxMonth) & "' and '" & MaxMonth & "' order by TrendMonth desc"

'strDetail = "select TrendMonth, ROs, InsuranceQ, QualityQ, Comm, OnTime, ReturnVisit, (avg(RecShopNum)) As RecShopNum " & _
'		" from tblCollisionTrend where (ShopID=" & ishopid & ") and TrendMonth between '" & dateadd("M",-11,MaxMonth) & "' and '" & MaxMonth & "'" & _
'		" group by TrendMonth, ROs, InsuranceQ, QualityQ, Comm, OnTime, ReturnVisit, RecShop " & _
'		" order by TrendMonth desc"

rsDetail.open strDetail,cnlist




if rsdetail.eof then
	response.write "There are no completed months to display."
	response.write "<br>"
else
	Response.Write "<br>"
	Response.Write "<b>Note:&nbsp;</b>"
	Response.Write "<font color=blue>CSi percentage for months that still have pending ROs are based on surveys completed to date." & "<br>"
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;These results may not be the final result for the month." & "<br>"
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please see the RO List to verify if there are any ROs pending." & "<br>"
	Response.Write "<br></font>"
	
	drptmonth = rsdetail.Fields("TrendMonth")
	response.write "<TABLE width=""800"">"
	response.write "<TR>"
	response.write "<TH align=left width=""80""></TH>"
	response.write "<TH align=left width=""80""><font size=2>ROs</TH>"
	response.write "<TH align=left width=""80""><font size=2>Refer Shop</TH>"
	response.write "<TH align=left width=""80""><font size=2>Insurance</TH>"
	response.write "<TH align=left width=""80""><font size=2>Quality</TH>"
	response.write "<TH align=left width=""80""><font size=2>Informed</TH>"
	response.write "<TH align=left width=""80""><font size=2>On Time</TH>"
	response.write "<TH align=left width=""80""><font size=2>Return Rate</TH>"
	response.write "<TH align=right width=""80""><font size=2></TH>"
	response.write "</tr>"
	do until rsdetail.eof
		Response.write "<tr>"
		strMonth = rsdetail.Fields("TrendMonth")
		stryear = year(rsdetail.Fields("TrendMonth"))
		stryear = right(stryear,2)
		strmonth=MonthName(DatePart("m", strmonth),True)
		strmonth = strmonth & " " & stryear
		response.write "<td><font size=2>" & strmonth & "</td>"
		response.write "<td align=left><font size=2><b>" & rsdetail.fields("ROs") & "</b></td>"
		If IsNumeric(rsdetail.fields("RecShopNum")) then 
			response.write "<td align=left><font size=2>" & round(rsdetail.fields("RecShopNum"),2) & "</td>"
		Else
			response.write "<td align=left><font size=2>N/A</td>"
		End If
		If IsNumeric(rsdetail.fields("InsuranceQ")) then 
			response.write "<td align=left><font size=2>" & round(rsdetail.fields("InsuranceQ"),2) & "</td>"
		Else
			response.write "<td align=left><font size=2>N/A</td>"
		End If
		If IsNumeric(rsdetail.fields("QualityQ")) then 
			response.write "<td align=left><font size=2>" & round(rsdetail.fields("QualityQ"),2) & "</td>"
		Else
			response.write "<td align=left><font size=2>N/A</td>"
		End If
		If IsNumeric(rsdetail.fields("Comm")) then 
			response.write "<td align=left><font size=2>" & formatpercent(rsdetail.fields("Comm"),1) & "</td>"
		Else
			response.write "<td align=left><font size=2>N/A</td>"
		End If
		If IsNumeric(rsdetail.fields("OnTime")) then 
			response.write "<td align=left><font size=2>" & formatpercent(rsdetail.fields("OnTime"),1) & "</td>"
		Else
			response.write "<td align=left><font size=2>N/A</td>"
		End If
		If IsNumeric(rsdetail.fields("ReturnVisit")) then 
			response.write "<td align=left><font size=2>" & formatpercent(rsdetail.fields("ReturnVisit"),1) & "</td>"
		Else
			response.write "<td align=left><font size=2>N/A</td>"
		End If
		response.write "<td align=right><font size=2>" & strmonth & "</td>"
		Response.write "</tr>"
		rsdetail.movenext
	loop
	rsdetail.Close
	
'get the 12M number now....	
	strDetail = "select sum(ROs) as Records, (sum(ros*QualityQ))/(sum(ros)) as QualityQ, (sum(ros*InsuranceQ))/(sum(ros)) as InsuranceQ, " & _
		" (sum(ros*comm))/(sum(ros)) as Comm, (sum(ros*ontime))/(sum(ros)) as OnTime, (sum(ros*returnvisit))/(sum(ros)) as ReturnVisit, " & _
		" (sum(ros*RecShopNum))/(sum(ros)) as RecShopNum " & _
		" from tblCollisionTrend where (ShopID=" & ishopid & ") and TrendMonth between '" & dateadd("M",-11,MaxMonth) & "' and '" & MaxMonth & "'"
		
		rsDetail.open strDetail,cnlist
		
		Response.write "<tr>"
		response.write "<td><font size=2>Totals</td>"
		response.write "<td align=left><font size=2><b>" & rsdetail.fields("Records") & "</b></td>"
		response.write "<td align=left><font size=2>" & round(rsdetail.fields("RecShopNum"),2) & "</td>"
		
		If IsNumeric(rsdetail.fields("InsuranceQ")) then 
			response.write "<td align=left><font size=2>" & round(rsdetail.fields("InsuranceQ"),2) & "</td>"
		Else
			response.write "<td align=left><font size=2>N/A</td>"
		End If
		
		If IsNumeric(rsdetail.fields("QualityQ")) then 
			response.write "<td align=left><font size=2>" & round(rsdetail.fields("QualityQ"),2) & "</td>"
		Else
			response.write "<td align=left><font size=2>N/A</td>"
		End If
		
		response.write "<td align=left><font size=2>" & formatpercent(rsdetail.fields("Comm"),1) & "</td>"
		response.write "<td align=left><font size=2>" & formatpercent(rsdetail.fields("OnTime"),1) & "</td>"
		response.write "<td align=left><font size=2>" & formatpercent(rsdetail.fields("ReturnVisit"),1) & "</td>"
		response.write "<td align=right><font size=2>Totals</td>"
		Response.write "</tr>"
		response.write "</TABLE>"
end if



response.write "<TABLE width=""1100"" cellSpacing=1 cellPadding=1 border=0>"
Response.write "<tr>"
Response.write "<td align=center><font size=1><em>"
Response.Write "<p>"
response.write "Information compiled by CSi Complete, a national provider of CSI and call center services to business. For more information please call (800) 343-0641. Copyright 2012 CSi Complete</FONT></B></I></TD>"
response.write "</font></em>"
Response.Write "</td>"
Response.Write "</tr>"
response.write "</TABLE>"
response.write "</BODY></HTML>"
%>
