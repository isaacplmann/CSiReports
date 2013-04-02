<%@ Language=VBScript %>
<%Response.Buffer=True%>
<html>
<head>
<title>Service Writer Report</title>
</head>
<body><font face=arial size=2>
<%
dim dRptMonth
dim iParentID

dRptMonth=Request.QueryString("Month")
dRptMonth=dateadd("m",1,dRptMonth)
iparentid=session("id")

set cnList=server.createobject("adodb.Connection")
cnList.open session("Constring")

	Response.Write "<tr>"
	Response.Write "<td align=left><font size=3 color=brown><strong>"
	Response.Write session("sessionname")
	Response.Write "</strong></font></td>"
	Response.Write "</tr>"
	Response.Write "<br>"
	Response.Write "<br>"
	

	
		
strDetail="select * from (select tblCollisionLogons.Name, tblCollisionDetail.StatusID, tblCollisionLogons.ShopID, tblCollisionDetail.Estimator, Count(*) AS Records,  " & _
	"left(datename(" & "" & "m" & "" & ",tblCollisionDetail.RptMonth),3) + ' ' + right(convert(nvarchar,year(tblCollisionDetail.RptMonth)),2) as Period, " & _
	"Avg(Abs(tblCollisionDetailAll.BodyQ)) AS Body, Avg(Abs(tblCollisionDetailAll.PaintQ)) AS Paint, Avg(Abs(tblCollisionDetailAll.MechQ)) AS Mech, Avg(Abs(tblCollisionDetailAll.DetailQ)) AS Detail,  " & _
	"Avg(Abs(tblCollisionDetailAll.CommQ)) AS Comm, Avg(Abs(tblCollisionDetailAll.OnTimeQ)) AS OnTime, Avg(Abs(tblCollisionDetailAll.ReturnQ)) AS ReturnVisit, Avg(Abs(tblCollisionDetailAll.RecShopQ)) AS RecShop,  " & _
	"Avg(tblCollisionDetail.CSi) As CSi, Avg(convert(float(2),tblCollisionDetailAll.RecShopNum)) As RecShopNum, 0 as Sort " & _
	"FROM tblCollisionLogons INNER JOIN tblCollisionDetail INNER JOIN tblCollisionDetailAll ON tblCollisionDetail.JobID = tblCollisionDetailAll.JobID ON tblCollisionLogons.ShopID = tblCollisionDetail.ShopID " & _
	"Where tblCollisionDetail.StatusID=1 and RptMonth = '" & dRptMonth & "'" & _
	"and tblCollisionLogons.ParentID=" & iparentid & " group by tblCollisionLogons.Name, tblCollisionLogons.ShopID, tblCollisionDetail.Estimator, tblCollisionDetail.StatusID, tblCollisionDetail.RptMonth) C inner join "

	
strDetail= strDetail & "(select tblCollisionLogons.Name As Name12m, tblCollisionDetail.StatusID As StatusID12m, tblCollisionLogons.ShopID As ShopID12m, tblCollisionDetail.Estimator As Estimator12m, Count(*) AS Records12m, " & _
	"'12m' as Period12m, " & _
	"Avg(Abs(tblCollisionDetailAll.BodyQ)) AS Body12m, Avg(Abs(tblCollisionDetailAll.PaintQ)) AS Paint12m, Avg(Abs(tblCollisionDetailAll.MechQ)) AS Mech12m, Avg(Abs(tblCollisionDetailAll.DetailQ)) AS Detail12m,  " & _
	"Avg(Abs(tblCollisionDetailAll.CommQ)) AS Comm12m, Avg(Abs(tblCollisionDetailAll.OnTimeQ)) AS OnTime12m, Avg(Abs(tblCollisionDetailAll.ReturnQ)) AS ReturnVisit12m, Avg(Abs(tblCollisionDetailAll.RecShopQ)) AS RecShop12m,  " & _
	"Avg(tblCollisionDetail.CSi) As CSi12m, Avg(convert(float(2),tblCollisionDetailAll.RecShopNum)) As RecShopNum12m, 2 as Sort12m " & _
	"FROM tblCollisionLogons INNER JOIN tblCollisionDetail INNER JOIN tblCollisionDetailAll ON tblCollisionDetail.JobID = tblCollisionDetailAll.JobID ON tblCollisionLogons.ShopID = tblCollisionDetail.ShopID " & _
	"Where tblCollisionDetail.StatusID=1 and RptMonth between '" & dateadd("m",-11,dRptMonth) & "' and '" & dRptMonth & "'" & _
	"and tblCollisionLogons.ParentID=" & iparentid & " " & _
	"group by tblCollisionLogons.Name, tblCollisionLogons.ShopID, tblCollisionDetail.Estimator, tblCollisionDetail.StatusID, tblCollisionDetail.RptMonth) E on E.Estimator12m=C.Estimator order by Estimator, Sort "


set rsdetail=server.createobject("adodb.recordset")
rsDetail.open strDetail,cnlist

if rsdetail.eof then
	response.write "There are no completed months to display."
	response.write "<br>"
else
	Response.Write "<br>"
	
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
	response.write "<TH align=center width=""60""><font size=2>0-10</TH>"
	response.write "<TH align=right width=""60""><font size=2></TH>"
	response.write "</tr>"
	do until rsdetail.eof
		Response.write "<tr>"
		response.write "<td colspan=7><font size=2 color=blue><B>" & rsdetail.fields("Estimator") & "</b></td>"
		response.write "<td colspan=7 align=right><font size=2 color=blue><B>" & rsdetail.fields("Estimator") & "</b></td>"
		Response.Write "</tr>"
		Response.write "<tr>"
		Response.Write "<td align=center><font size=2><b>" & rsdetail.Fields("Period") & "</b></td>"
		response.write "<td align=center><font size=2><b>" & rsdetail.fields("Records") & "</b></td>"
		If IsNumeric(rsdetail.fields("Body")) then 
			response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Body"),1) & "</td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		If IsNumeric(rsdetail.fields("Paint")) then 
			response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Paint"),1) & "</td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		If IsNumeric(rsdetail.fields("Mech")) then 
			response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Mech"),1) & "</td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		If IsNumeric(rsdetail.fields("Detail")) then 
			response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Detail"),1) & "</td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		If IsNumeric(rsdetail.fields("Comm")) then 
			response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Comm"),1) & "</td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		If IsNumeric(rsdetail.fields("OnTime")) then 
			response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("OnTime"),1) & "</td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		If IsNumeric(rsdetail.fields("ReturnVisit")) then 
			response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("ReturnVisit"),1) & "</td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		If IsNumeric(rsdetail.fields("RecShop")) then 
			response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("RecShop"),1) & "</td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		If IsNumeric(rsdetail.fields("CSi")) then 
			response.write "<td align=center><font size=2 color=red>" & formatpercent(rsdetail.fields("CSi"),1) & "</td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		If IsNumeric(rsdetail.fields("RecShopNum")) then 
			response.write "<td align=center><font size=2><B>" & formatnumber(rsdetail.fields("RecShopNum"),2) & "</b></td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		response.write "<td align=right><font size=2><B>" & rsdetail.fields("Period") & "</b></td>"
		Response.write "</tr>"
		
		Response.write "<tr>"
		Response.Write "<td align=center><font size=2><b>" & rsdetail.Fields("Period12m") & "</b></td>"
		response.write "<td align=center><font size=2><b>" & rsdetail.fields("Records12m") & "</b></td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Body12m"),1) & "</td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Paint12m"),1) & "</td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Mech12m"),1) & "</td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Detail12m"),1) & "</td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Comm12m"),1) & "</td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("OnTime12m"),1) & "</td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("ReturnVisit12m"),1) & "</td>"
		response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("RecShop12m"),1) & "</td>"
		response.write "<td align=center><font size=2 color=red>" & formatpercent(rsdetail.fields("CSi12m"),1) & "</td>"
		response.write "<td align=center><font size=2><B>" & formatnumber(rsdetail.fields("RecShopNum12m"),2) & "</b></td>"
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
response.write "Information compiled by CSi Complete, a national provider of CSI and call center services to business. For more information please call (800) 343-0641. Copyright 2010 CSi Complete</FONT></B></I></TD>"
response.write "</font></em>"
Response.Write "</td>"
Response.Write "</tr>"
response.write "</TABLE>"
response.write "</BODY></HTML>"
%>
