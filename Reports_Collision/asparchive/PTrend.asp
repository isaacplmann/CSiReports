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
	" (sum((convert(decimal(9,4),greeted)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as Greeted, " & _
	" (sum((convert(decimal(9,4),body)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as Body, " & _
	" (sum((convert(decimal(9,4),paint)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as Paint, " & _
	" (sum((convert(decimal(9,4),mech)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as Mech, " & _
	" (sum((convert(decimal(9,4),detail)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as Detail, " & _
	" (sum((convert(decimal(9,4),clean)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as Clean, " & _
	" (sum((convert(decimal(9,4),service)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as Service, " & _
	" (sum((convert(decimal(9,4),comm)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as Comm, " & _
	" (sum((convert(decimal(9,4),ontime)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as ontime, " & _
	" (sum((convert(decimal(9,4),returnvisit)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as returnvisit, " & _ 
	" (sum((convert(decimal(9,4),recshop)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as recshop, " & _
	" (sum((convert(decimal(9,4),handleclaim)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as handleclaim, " & _
	" (sum((convert(decimal(9,4),recins)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as recins, " & _
	" (sum((convert(decimal(9,4),rental)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as rental, " & _
	" (sum((convert(decimal(9,4),CSi)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as CSi, " & _
	" TrendMonth " & _
	" FROM tblCollisionTrend as T " & _
	" INNER JOIN tblCollisionLogons as L ON T.ShopID = L.ShopID " & _
	" WHERE (L.ParentID=" & iparent & ")" & _
	" and TrendMonth between '" & dateadd("M",-11,MaxMonth) & "' and '" & MaxMonth & "'" & _
	" Group By TrendMonth"  & _
	" order by trendmonth desc"
	


'strdetail= "SELECT D.RptMonth, Count(D.JobID) AS ROS, Avg([GreetedQ]*-1) AS Greeted, Avg([BodyQ]*-1) AS Body, Avg([PaintQ]*-1) AS Paint, " & _
'	" Avg([MechQ]*-1) AS Mech, Avg([DetailQ]*-1) AS Detail, Avg([CleanQ]*-1) AS Clean, Avg([ServiceQ]*-1) AS Service, " & _
'	" Avg([CommQ]*-1) AS Comm, Avg([OnTimeQ]*-1) AS OnTime, Avg([ReturnQ]*-1) AS ReturnVisit, Avg([RecShopQ]*-1) AS RecShop, " & _
'	" Avg([ClaimQ]*-1) AS HandleClaim, Avg([RecInsQ]*-1) AS RecIns, Avg([CSi]) AS CSi " & _
'	" FROM (tblCollisionDetail as D INNER JOIN tblCollisionDetailAll as DA ON D.JobID = DA.JobID) " &_
'	" INNER JOIN tblCollisionLogons as L ON D.ShopID = L.ShopID " & _
'	" WHERE (((L.ParentID)=" & session("id") & ") AND ((D.StatusID)=1)) " & _
'	" GROUP BY D.RptMonth " & _
'	" HAVING (((D.RptMonth) Between '" & dateadd("m",-11,drptmonth) & "' And '" & drptmonth & "'))" & _
'	" ORDER BY RptMonth desc"


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
	response.write "<TH align=center width=""60""><font size=2>Greet</TH>"
	response.write "<TH align=center width=""60""><font size=2>Body</TH>"
	response.write "<TH align=center width=""60""><font size=2>Paint</TH>"
	response.write "<TH align=center width=""60""><font size=2>Mech</TH>"
	response.write "<TH align=center width=""60""><font size=2>Detail</TH>"
	response.write "<TH align=center width=""60""><font size=2>Clean</TH>"
	response.write "<TH align=center width=""60""><font size=2>Service</TH>"
	response.write "<TH align=center width=""60""><font size=2>Comm</TH>"
	response.write "<TH align=center width=""60""><font size=2>Time</TH>"
	response.write "<TH align=center width=""70""><font size=2>2nd Visit</TH>"
	response.write "<TH align=center width=""60""><font size=2>RecShop</TH>"
	response.write "<TH align=center width=""60""><font size=2>Claim</TH>"
	response.write "<TH align=center width=""60""><font size=2>Rec Ins</TH>"
	response.write "<TH align=center width=""60""><font size=2 color=red>CSi</TH>"
	response.write "<TH align=right width=""60""><font size=2></TH>"
	response.write "</tr>"

	response.write "<tr>"
	Response.Write "<td colspan=16>"
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
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("Greeted"),4)  & "%</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("Body"),4) & "%</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("Paint"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("Mech"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("Detail"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("Clean"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("Service"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("Comm"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("OnTime"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("ReturnVisit"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("RecShop"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("HandleClaim"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("RecIns"),4) & "</td>"
		response.write "<td align=center><font size=2 color=red>" & left(rsdetail.fields("CSi"),4) & "</td>"
		response.write "<td align=right><font size=2>" & strmonth & "</td>"
		Response.write "</tr>"
		rsdetail.movenext
	loop
	rsdetail.Close
	Response.Flush
'get the 12M number now....	

strdetail = "select sum(ROs) as ROs, " & _
	" (sum((convert(decimal(9,4),greeted)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as Greeted, " & _
	" (sum((convert(decimal(9,4),body)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as Body, " & _
	" (sum((convert(decimal(9,4),paint)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as Paint, " & _
	" (sum((convert(decimal(9,4),mech)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as Mech, " & _
	" (sum((convert(decimal(9,4),detail)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as Detail, " & _
	" (sum((convert(decimal(9,4),clean)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as Clean, " & _
	" (sum((convert(decimal(9,4),service)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as Service, " & _
	" (sum((convert(decimal(9,4),comm)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as Comm, " & _
	" (sum((convert(decimal(9,4),ontime)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as ontime, " & _
	" (sum((convert(decimal(9,4),returnvisit)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as returnvisit, " & _ 
	" (sum((convert(decimal(9,4),recshop)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as recshop, " & _
	" (sum((convert(decimal(9,4),handleclaim)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as handleclaim, " & _
	" (sum((convert(decimal(9,4),recins)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as recins, " & _
	" (sum((convert(decimal(9,4),rental)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as rental, " & _
	" (sum((convert(decimal(9,4),CSi)* Convert(Decimal(9,2),[ROS])))/(sum(ROs))*100) as CSi " & _
	" FROM tblCollisionTrend as T " & _
	" INNER JOIN tblCollisionLogons as L ON T.ShopID = L.ShopID " & _
	" WHERE (L.ParentID=" & iparent & ")" & _
	" and TrendMonth between '" & dateadd("M",-11,MaxMonth) & "' and '" & MaxMonth & "'" 

	response.write "<tr>"
	Response.Write "<td colspan=16>"
	'Response.Write strdetail
	Response.Write "</td>"
	response.write "</tr>"


		rsDetail.open strDetail,cnlist
do until rsdetail.EOF		
		Response.write "<tr>"
		response.write "<td><font size=2>Totals</td>"
		response.write "<td align=center><font size=2><b>" & rsdetail.fields("ROs") & "</b></td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("Greeted"),4)  & "%</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("Body"),4) & "%</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("Paint"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("Mech"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("Detail"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("Clean"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("Service"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("Comm"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("OnTime"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("ReturnVisit"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("RecShop"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("HandleClaim"),4) & "</td>"
		response.write "<td align=center><font size=2>" & left(rsdetail.fields("RecIns"),4) & "</td>"
		response.write "<td align=center><font size=2 color=red>" & left(rsdetail.fields("CSi"),4) & "</td>"
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
response.write "Information compiled by CSi Complete, a national provider of CSI and call center services to business. For more information please call (800) 343-0641. Copyright 2003 CSi Complete</FONT></B></I></TD>"
response.write "</font></em>"
Response.Write "</td>"
Response.Write "</tr>"
response.write "</TABLE>"
response.write "</BODY></HTML>"
%>
