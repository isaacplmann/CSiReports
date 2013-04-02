<%@ Language=VBScript %>
<%Response.Buffer=True%>
<HTML>
<HEAD>
<TITLE>Trend Report</TITLE>
</HEAD>
<BODY><font face=arial size=2>
<%
dim iShopID 
dim dRptMonth
dim strMonth
dim strYear

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

strDetail = "select * from tblCollisionTrendToy where ShopID=" & ishopid & " order by TrendMonth desc"
set rsdetail=server.createobject("adodb.recordset")
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
	response.write "<TABLE>"
	response.write "<TR>"
	response.write "<TH align=left width=""60""></TH>"
	response.write "<TH align=center width=""50""><font size=2>ROs</TH>"
	response.write "<TH align=center width=""100""><font size=2>Quality</TH>"
	response.write "<TH align=center width=""100""><font size=2>Treated</TH>"
	response.write "<TH align=center width=""100""><font size=2>Ready</TH>"
	response.write "<TH align=center width=""100""><font size=2>Refer Shop</TH>"
	response.write "<TH align=center width=""100""><font size=2 color=red>CSi</TH>"
	response.write "</tr>"
	do until rsdetail.eof
		Response.write "<tr>"
		strMonth = rsdetail.Fields("TrendMonth")
		stryear = year(rsdetail.Fields("TrendMonth"))
		stryear = right(stryear,2)
		strmonth=MonthName(DatePart("m", strmonth),True)
		strmonth = strmonth & " " & stryear
		response.write "<td><font size=2>" & strmonth & "</td>"
		response.write "<td align=center><font size=2><b>" & rsdetail.fields("ROs") & "</b></td>"
		response.write "<td align=center><font size=2>" & formatnumber(rsdetail.fields("Quality"),2) & "</td>"
		response.write "<td align=center><font size=2>" & formatnumber(rsdetail.fields("Treated"),2) & "</td>"
		response.write "<td align=center><font size=2>" & formatnumber(rsdetail.fields("OnTime"),2) & "</td>"
		response.write "<td align=center><font size=2>" & formatnumber(rsdetail.fields("RecShop"),2) & "</td>"
		response.write "<td align=center><font size=2 color=red>" & formatpercent(rsdetail.fields("CSi"),1) & "</td>"
		response.write "<td align=right><font size=2>" & strmonth & "</td>"
		Response.write "</tr>"
		rsdetail.movenext
	loop
	rsdetail.Close
	
'get the 12M number now....	
	strDetail = "select sum(ROs) as Records, (sum(ros*quality))/(sum(ros)) as Body, " & _
		" (sum(ros*treated))/(sum(ros)) as Service, " & _
		" (sum(ros*ontime))/(sum(ros)) as OnTime, (sum(ros*recshop))/(sum(ros)) as RecShop, " & _
		" (sum(ros*csi))/(sum(ros)) as CSi " & _
		" from tblCollisionTrendToy where (ShopID=" & ishopid & ") and " & _
		" (TrendMonth between'" & dateadd("m",-11,drptmonth) & "' and '" & drptmonth & "')"
		
		rsDetail.open strDetail,cnlist
		
		
		
		Response.write "<tr>"
		response.write "<td><font size=2>12M</td>"
		response.write "<td align=center><font size=2><b>" & rsdetail.fields("Records") & "</b></td>"
		response.write "<td align=center><font size=2>" & formatnumber(rsdetail.fields("Body"),2) & "</td>"
		response.write "<td align=center><font size=2>" & formatnumber(rsdetail.fields("Service"),2) & "</td>"
		response.write "<td align=center><font size=2>" & formatnumber(rsdetail.fields("OnTime"),2) & "</td>"
		response.write "<td align=center><font size=2>" & formatnumber(rsdetail.fields("RecShop"),2) & "</td>"
		response.write "<td align=center><font size=2 color=red>" & formatpercent(rsdetail.fields("CSi"),1) & "</td>"
		response.write "<td align=right><font size=2>12M</td>"
		Response.write "</tr>"
	response.write "</TABLE>"
end if

'Response.Write strdetail & "<br>"


Response.Write "<p>"
response.write "<TABLE cellSpacing=1 cellPadding=1 border=0>"
Response.write "<tr>"
Response.write "<td align=center><font size=1><em>"
response.write "Information compiled by CSi Complete, a national provider of CSI and call center services to business. " & "</td></tr>"
Response.write "<tr>"
Response.write "<td align=center><font size=1><em>"
Response.write "For more information please call (800) 343-0641. Copyright 2003 CSi Complete</FONT></B></I></TD>"
response.write "</font></em>"
Response.Write "</tr>"
response.write "</TABLE>"
response.write "</BODY></HTML>"
%>
