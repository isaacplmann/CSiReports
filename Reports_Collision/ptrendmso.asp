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
dim iParent
ishopid=Request.QueryString("shopid")
dRptMonth=Request.QueryString("month")
set cnList=server.createobject("adodb.Connection")
cnList.open session("Constring")

'get the parentID to find this stuff out!!!
iparent= session("id")



set rsdetail=server.createobject("adodb.recordset")
'Get the latest trend month
strDetail = "select max(trendmonth) as MaxMonth from tblCollisionTrend T inner join tblcollisionlogons L on T.ShopID = L.ShopID where ParentID=" & iparent
rsDetail.open strDetail, cnlist
if rsdetail.eof then 
	MaxMonth = date()
else
	MaxMonth = rsdetail.fields("MaxMonth")
end if
rsDetail.close



'THIS NEEDS TO BE LIKE THE 12M NUMBERS
strdetail = "select sum(ROs) as ROs, " & _
	" (sum((convert(decimal(9,4),body)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as Body, " & _
	" (sum((convert(decimal(9,4),paint)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as Paint, " & _
	" (sum((convert(decimal(9,4),mech)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as Mech, " & _
	" (sum((convert(decimal(9,4),detail)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as Detail, " & _
	" (sum((convert(decimal(9,4),comm)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as Comm, " & _
	" (sum((convert(decimal(9,4),ontime)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as ontime, " & _
	" (sum((convert(decimal(9,4),returnvisit)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as returnvisit, " & _ 
	" (sum((convert(decimal(9,4),recshop)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as recshop, " & _
	" (sum((convert(decimal(9,4),CSi)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as CSi, " & _
	" (sum(convert(float(2),RecShopNum)*[ROS])/(sum([ROs]))) as recshopnum, " & _
	" TrendMonth " & _
	" FROM tblCollisionTrend as T " & _
	" INNER JOIN tblCollisionLogons as L ON T.ShopID = L.ShopID " & _
	" WHERE (L.ParentID=" & iparent & ")" & _
	" and TrendMonth between '" & dateadd("M",-11,MaxMonth) & "' and '" & MaxMonth & "'" & _
	" Group By TrendMonth"  & _
	" order by trendmonth desc"

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
	
	'drptmonth = rsdetail.Fields("TrendMonth")
	response.write "<TABLE width=""1100"">"
	response.write "<TR>"
	response.write "<TH align=left width=""60""></TH>"
	response.write "<TH align=center width=""60""><font size=2>ROs</TH>"
	response.write "<TH align=center width=""60""><font size=2>Body</TH>"
	response.write "<TH align=center width=""60""><font size=2>Paint</TH>"
	response.write "<TH align=center width=""60""><font size=2>Mech</TH>"
	response.write "<TH align=center width=""60""><font size=2>Detail</TH>"
	response.write "<TH align=center width=""60""><font size=2>Comm</TH>"
	response.write "<TH align=center width=""60""><font size=2>Time</TH>"
	response.write "<TH align=center width=""70""><font size=2>2nd Visit</TH>"
	response.write "<TH align=center width=""60""><font size=2>RecShop</TH>"
	response.write "<TH align=center width=""60""><font size=2 color=red>CSi</TH>"
	response.write "<TH align=center width=""60""><font size=2>0 - 10</TH>"
	response.write "<TH align=right width=""60""><font size=2></TH>"
	response.write "</tr>"

	response.write "<tr>"
	Response.Write "<td colspan=11>"
	'Response.Write strdetail
	Response.Write "</td>"
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
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("Body"),4) & "%</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("Paint"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("Mech"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("Detail"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("Comm"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("OnTime"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("ReturnVisit"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("RecShop"),4) & "</td>"
		response.write "<td align=center><font size=2 color=red>" & left(rsdetail.fields("CSi"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("RecShopNum"),4) & "</td>"
		response.write "<td align=right><font size=2>" & strmonth & "</td>"
		Response.write "</tr>"
		rsdetail.movenext
	loop
	rsdetail.Close
	Response.Flush
'get the 12M number now....	

strdetail = "select sum(ROs) as ROs, " & _
	" (sum((convert(decimal(9,4),body)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as Body, " & _
	" (sum((convert(decimal(9,4),paint)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as Paint, " & _
	" (sum((convert(decimal(9,4),mech)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as Mech, " & _
	" (sum((convert(decimal(9,4),detail)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as Detail, " & _
	" (sum((convert(decimal(9,4),comm)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as Comm, " & _
	" (sum((convert(decimal(9,4),ontime)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as ontime, " & _
	" (sum((convert(decimal(9,4),returnvisit)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as returnvisit, " & _ 
	" (sum((convert(decimal(9,4),recshop)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as recshop, " & _
	" (sum((convert(decimal(9,4),CSi)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as CSi, " & _
	" (sum(convert(float(2),RecShopNum)*[ROS])/(sum([ROs]))) as recshopnum " & _
	" FROM tblCollisionTrend as T " & _
	" INNER JOIN tblCollisionLogons as L ON T.ShopID = L.ShopID " & _
	" WHERE (L.ParentID=" & iparent & ")" & _
	" and TrendMonth between '" & dateadd("M",-11,MaxMonth) & "' and '" & MaxMonth & "'" 

	response.write "<tr>"
	Response.Write "<td colspan=11>"
	'Response.Write strdetail
	Response.Write "</td>"
	response.write "</tr>"


		rsDetail.open strDetail,cnlist
do until rsdetail.EOF		
		Response.write "<tr>"
		response.write "<td><font size=2>Totals</td>"
		response.write "<td align=center><font size=2><b>" & rsdetail.fields("ROs") & "</b></td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("Body"),4) & "%</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("Paint"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("Mech"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("Detail"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("Comm"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("OnTime"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("ReturnVisit"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("RecShop"),4) & "</td>"
		response.write "<td align=center><font size=2 color=red>" & left(rsdetail.fields("CSi"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("RecShopNum"),4) & "</td>"
		response.write "<td align=right><font size=2>Totals</td>"
		Response.write "</tr>"
	rsdetail.MoveNext
loop
	response.write "</TABLE>"
end if



response.write "<TABLE width=""1100"" cellSpacing=1 cellPadding=1 border=0>"
Response.write "<tr>"
Response.write "<td align=center><font size=1><em>"
Response.Write "<p>"
response.write "Information compiled by CSi Complete, a national provider of CSI and call center services to business. For more information please call (800) 343-0641. Copyright 2010 CSi Complete</FONT></B></I></TD>"
response.write "</font></em>"
Response.Write "</td>"
Response.Write "</tr>"
response.write "</TABLE>"
response.write "</BODY></HTML>"
%>
