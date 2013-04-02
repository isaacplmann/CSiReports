<%@ Language=VBScript %>
<%Response.Buffer=True%>
<HTML>
<HEAD>
<TITLE>Trend Report</TITLE>
</HEAD>
<BODY><font face=arial size=2>
<%
dim dRptMonth
dim iParentID

dRptMonth=Request.QueryString("Month")
iparentid=session("id")

set cnList=server.createobject("adodb.Connection")
cnList.open session("Constring")

	Response.Write "<TR>"
	Response.Write "<Td align=left><font size=3 color=brown><strong>"
	Response.Write session("sessionname")
	Response.Write "</strong></font></Td>"
	Response.Write "</TR>"
	Response.Write "<br>"
	Response.Write "<br>"

strDetail="select * from (select tblCollisionLogons.Name, tblCollisionTrend.ShopID, ROs as Records, " & _
	"RecShopNum, Body, Paint, Mech, Detail, Comm, OnTime, ReturnVisit, RecShop, " & _
	" HandleClaim, CSi, left(datename(" & "" & "m" & "" & ",TrendMonth),3) + ' ' + right(convert(nvarchar,year(TrendMonth)),2) as Period, " & _
	" 0 as Sort from tblCollisionTrend inner join tblCollisionLogons " & _
	" on tblCollisionTrend.ShopID=tblCollisionLogons.ShopID where TrendMonth = '" & drptmonth & "'" & _
	" and tblCollisionLogons.ParentID=" & iparentid & ") C inner join "

strDetail = strdetail & "(select tblCollisionLogons.Name as Name12m, tblCollisionTrend.ShopID as ShopID12m, sum(ROs) as Records12m, " & _
	" avg(RecShopNum) as RecShopNum12m, (sum(ros*body))/(sum(ros)) as Body12m, " & _
	" (sum(ros*paint))/(sum(ros)) as Paint12m, (sum(ros*mech))/(sum(ros)) as Mech12m, " & _
	" (sum(ros*detail))/(sum(ros)) as Detail12m, " & _
	" (sum(ros*comm))/(sum(ros)) as Comm12m, " & _
	" (sum(ros*ontime))/(sum(ros)) as OnTime12m, (sum(ros*returnvisit))/(sum(ros)) as ReturnVisit12m, " & _
	" (sum(ros*recshop))/(sum(ros)) as RecShop12m, (sum(ros*handleclaim))/(sum(ros)) as HandleClaim12m, " & _
	" (sum(ros*recins))/(sum(ros)) as RecIns12m, (sum(ros*csi))/(sum(ros)) as CSi12m, '12M' as Period12m, 1 as Sort12m " & _
	" from tblCollisionTrend inner join tblCollisionLogons on tblCollisionTrend.ShopID=tblCollisionLogons.ShopID " & _
	" where TrendMonth between '" & dateadd("m",-11,dRptMonth) & "' and '" & dRptMonth & "' and tblCollisionLogons.ParentID=" & iparentid  & " " & _
	" group by tblCollisionTrend.ShopID,tblCollisionLogons.Name) M on M.ShopID12m=C.ShopID order by Name, Sort"

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
	
	response.write "<TABLE width=""1100"">"
	response.write "<TR>"
	response.write "<TH align=left width=""60""></TH>"
	response.write "<TH align=center width=""60""><font size=2>ROs</TH>"
	response.write "<TH align=center width=""60""><font size=2>Shop Rating</TH>"
	response.write "<TH align=center width=""60""><font size=2>Body</TH>"
	response.write "<TH align=center width=""60""><font size=2>Paint</TH>"
	response.write "<TH align=center width=""60""><font size=2>Mech</TH>"
	response.write "<TH align=center width=""60""><font size=2>Detail</TH>"
	response.write "<TH align=center width=""60""><font size=2>Informed</TH>"
	response.write "<TH align=center width=""60""><font size=2>On Time</TH>"
	response.write "<TH align=center width=""70""><font size=2>2nd Visit</TH>"
	response.write "<TH align=center width=""60""><font size=2>RecShop</TH>"
	response.write "<TH align=center width=""60""><font size=2>Claim</TH>"
	response.write "<TH align=center width=""60""><font size=2 color=red>CSi</TH>"
	response.write "<TH align=right width=""60""><font size=2></TH>"
	response.write "</tr>"
	do until rsdetail.eof
		Response.write "<tr>"
		response.write "<td colspan=8><font size=2 color=blue><B>" & rsdetail.fields("Name") & "</b></td>"
		response.write "<td colspan=9 align=right><font size=2 color=blue><B>" & rsdetail.fields("Name") & "</b></td>"
		Response.Write "</tr>"
		Response.write "<tr>"
		Response.Write "<td align=center><font size=2><b>" & rsdetail.Fields("Period") & "</b></td>"
		response.write "<td align=center><font size=2><b>" & rsdetail.fields("Records") & "</b></td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("RecShopNum"),1) & "</td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Body"),1) & "</td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Paint"),1) & "</td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Mech"),1) & "</td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Detail"),1) & "</td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Comm"),1) & "</td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("OnTime"),1) & "</td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("ReturnVisit"),1) & "</td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("RecShop"),1) & "</td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("HandleClaim"),1) & "</td>"
		response.write "<td align=center><font size=2 color=red>" & formatpercent(rsdetail.fields("CSi"),1) & "</td>"
		response.write "<td align=right><font size=2><B>" & rsdetail.fields("Period") & "</b></td>"
		Response.write "</tr>"
		Response.write "<tr>"
		Response.Write "<td align=center><font size=2><b>" & rsdetail.Fields("Period12m") & "</b></td>"
		response.write "<td align=center><font size=2><b>" & rsdetail.fields("Records12m") & "</b></td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("RecShop12m"),1) & "</td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Body12m"),1) & "</td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Paint12m"),1) & "</td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Mech12m"),1) & "</td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Detail12m"),1) & "</td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Comm12m"),1) & "</td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("OnTime12m"),1) & "</td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("ReturnVisit12m"),1) & "</td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("RecShop12m"),1) & "</td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("HandleClaim12m"),1) & "</td>"
		response.write "<td align=center><font size=2 color=red>" & formatpercent(rsdetail.fields("CSi12m"),1) & "</td>"
		response.write "<td align=right><font size=2>" & rsdetail.Fields("Period12m") & "</td>"
		Response.write "</tr>"
		
		rsdetail.movenext
	loop
	rsdetail.Close
	response.write "</TABLE>"
end if

Response.Write "<p></p>"

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
