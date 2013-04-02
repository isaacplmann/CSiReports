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
'dRptMonth=dateadd("m",3,dRptMonth)
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
	
strDetail = "SELECT * FROM (SELECT tblCollisionDetail_GandC.Estimator, Count(*) AS Records,  " & _
	"Avg(Abs(tblCollisionDetailAll_GandC.BodyQ)) AS Body, Avg(Abs(tblCollisionDetailAll_GandC.PaintQ)) AS Paint, Avg(Abs(tblCollisionDetailAll_GandC.MechQ)) AS Mech, Avg(Abs(tblCollisionDetailAll_GandC.DetailQ)) AS Detail,  " & _
	"Avg(Abs(tblCollisionDetailAll_GandC.CommQ)) AS Comm, Avg(Abs(tblCollisionDetailAll_GandC.OnTimeQ)) AS OnTime, Avg(Abs(tblCollisionDetailAll_GandC.ReturnQ)) AS ReturnVisit, Avg(Abs(tblCollisionDetailAll_GandC.RecShopQ)) AS RecShop,  " & _
	"Avg(tblCollisionDetail_GandC.CSi) As CSi, Avg(convert(float(2),tblCollisionDetailAll_GandC.RecShopNum)) As RecShopNum, 0 as Sort, " & _
	"left(datename(" & "" & "m" & "" & ",tblCollisionDetail_GandC.RptMonth),3) + ' ' + right(convert(nvarchar,year(tblCollisionDetail_GandC.RptMonth)),2) as Period " & _
	"FROM tblCollisionLogons LEFT JOIN (tblCollisionDetail_GandC LEFT JOIN tblCollisionDetailAll_GandC ON tblCollisionDetail_GandC.JobID = tblCollisionDetailAll_GandC.JobID) ON tblCollisionLogons.ShopID = tblCollisionDetail_GandC.ShopID " & _
	"WHERE tblCollisionDetail_GandC.StatusID=1 and tblCollisionDetail_GandC.Estimator<>'' and tblCollisionDetail_GandC.RptMonth = '" & dRptMonth & "'" & _
	"GROUP BY tblCollisionDetail_GandC.Estimator, tblCollisionDetail_GandC.StatusID, tblCollisionDetail_GandC.RptMonth) C LEFT JOIN "
	
strDetail = strDetail & "(SELECT DISTINCT tblCollisionDetail_GandC.Estimator, Count(*) AS Records1m,  " & _
	"Avg(Abs(tblCollisionDetailAll_GandC.BodyQ)) AS Body1m, Avg(Abs(tblCollisionDetailAll_GandC.PaintQ)) AS Paint1m, Avg(Abs(tblCollisionDetailAll_GandC.MechQ)) AS Mech1m, Avg(Abs(tblCollisionDetailAll_GandC.DetailQ)) AS Detail1m,  " & _
	"Avg(Abs(tblCollisionDetailAll_GandC.CommQ)) AS Comm1m, Avg(Abs(tblCollisionDetailAll_GandC.OnTimeQ)) AS OnTime1m, Avg(Abs(tblCollisionDetailAll_GandC.ReturnQ)) AS ReturnVisit1m, Avg(Abs(tblCollisionDetailAll_GandC.RecShopQ)) AS RecShop1m,  " & _
	"Avg(tblCollisionDetail_GandC.CSi) As CSi1m, Avg(convert(float(2),tblCollisionDetailAll_GandC.RecShopNum)) As RecShopNum1m, 1 as Sort, " & _
	"'3m' as Period1m " & _
	"FROM tblCollisionLogons LEFT JOIN (tblCollisionDetail_GandC LEFT JOIN tblCollisionDetailAll_GandC ON tblCollisionDetail_GandC.JobID = tblCollisionDetailAll_GandC.JobID) ON tblCollisionLogons.ShopID = tblCollisionDetail_GandC.ShopID " & _
	"WHERE tblCollisionDetail_GandC.StatusID=1 and tblCollisionDetail_GandC.Estimator<>'' and tblCollisionDetail_GandC.RptMonth between '" & dateadd("m",-2,dRptMonth) & "' and '" & dRptMonth & "'" & _
	"GROUP BY tblCollisionDetail_GandC.Estimator, tblCollisionDetail_GandC.StatusID) D on D.Estimator=C.Estimator LEFT JOIN"

strDetail = strDetail & "(SELECT DISTINCT tblCollisionDetail_GandC.Estimator, Count(*) AS Records12m,  " & _
	"Avg(Abs(tblCollisionDetailAll_GandC.BodyQ)) AS Body12m, Avg(Abs(tblCollisionDetailAll_GandC.PaintQ)) AS Paint12m, Avg(Abs(tblCollisionDetailAll_GandC.MechQ)) AS Mech12m, Avg(Abs(tblCollisionDetailAll_GandC.DetailQ)) AS Detail12m,  " & _
	"Avg(Abs(tblCollisionDetailAll_GandC.CommQ)) AS Comm12m, Avg(Abs(tblCollisionDetailAll_GandC.OnTimeQ)) AS OnTime12m, Avg(Abs(tblCollisionDetailAll_GandC.ReturnQ)) AS ReturnVisit12m, Avg(Abs(tblCollisionDetailAll_GandC.RecShopQ)) AS RecShop12m,  " & _
	"Avg(tblCollisionDetail_GandC.CSi) As CSi12m, Avg(convert(float(2),tblCollisionDetailAll_GandC.RecShopNum)) As RecShopNum12m, 1 as Sort, " & _
	"'12m' as Period12m " & _
	"FROM tblCollisionLogons LEFT JOIN (tblCollisionDetail_GandC LEFT JOIN tblCollisionDetailAll_GandC ON tblCollisionDetail_GandC.JobID = tblCollisionDetailAll_GandC.JobID) ON tblCollisionLogons.ShopID = tblCollisionDetail_GandC.ShopID " & _
	"WHERE tblCollisionDetail_GandC.StatusID=1 and tblCollisionDetail_GandC.Estimator<>'' and tblCollisionDetail_GandC.RptMonth between '" & dateadd("m",-11,dRptMonth) & "' and '" & dRptMonth & "'" & _
	"GROUP BY tblCollisionDetail_GandC.Estimator, tblCollisionDetail_GandC.StatusID) E on E.Estimator=D.Estimator order by E.Estimator"




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
		Response.Write "<td align=center><font size=2><b>" & rsdetail.Fields("Period1m") & "</b></td>"
		response.write "<td align=center><font size=2><b>" & rsdetail.fields("Records1m") & "</b></td>"
		If IsNumeric(rsdetail.fields("Body1m")) then 
			response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Body1m"),1) & "</td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		If IsNumeric(rsdetail.fields("Paint1m")) then 
			response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Paint1m"),1) & "</td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		If IsNumeric(rsdetail.fields("Mech1m")) then 
			response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Mech1m"),1) & "</td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		If IsNumeric(rsdetail.fields("Detail1m")) then 
			response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Detail1m"),1) & "</td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		If IsNumeric(rsdetail.fields("Comm1m")) then 
			response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Comm1m"),1) & "</td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		If IsNumeric(rsdetail.fields("OnTime1m")) then 
			response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("OnTime1m"),1) & "</td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		If IsNumeric(rsdetail.fields("ReturnVisit1m")) then 
			response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("ReturnVisit1m"),1) & "</td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		If IsNumeric(rsdetail.fields("RecShop1m")) then 
			response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("RecShop1m"),1) & "</td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		If IsNumeric(rsdetail.fields("CSi1m")) then 
			response.write "<td align=center><font size=2 color=red>" & formatpercent(rsdetail.fields("CSi1m"),1) & "</td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		If IsNumeric(rsdetail.fields("RecShopNum1m")) then 
			response.write "<td align=center><font size=2><B>" & formatnumber(rsdetail.fields("RecShopNum1m"),2) & "</b></td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		response.write "<td align=right><font size=2><B>" & rsdetail.fields("Period1m") & "</b></td>"
		Response.write "</tr>"
		
		Response.write "<tr>"
		Response.Write "<td align=center><font size=2><b>" & rsdetail.Fields("Period12m") & "</b></td>"
		response.write "<td align=center><font size=2><b>" & rsdetail.fields("Records12m") & "</b></td>"
		If IsNumeric(rsdetail.fields("Body12m")) then 
			response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Body12m"),1) & "</td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		If IsNumeric(rsdetail.fields("Paint12m")) then 
			response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Paint12m"),1) & "</td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		If IsNumeric(rsdetail.fields("Mech12m")) then 
			response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Mech12m"),1) & "</td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		If IsNumeric(rsdetail.fields("Detail12m")) then 
			response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Detail12m"),1) & "</td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		If IsNumeric(rsdetail.fields("Comm12m")) then 
			response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("Comm12m"),1) & "</td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		If IsNumeric(rsdetail.fields("OnTime12m")) then 
			response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("OnTime12m"),1) & "</td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		If IsNumeric(rsdetail.fields("ReturnVisit12m")) then 
			response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("ReturnVisit12m"),1) & "</td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		If IsNumeric(rsdetail.fields("RecShop12m")) then 
			response.write "<td align=center><font size=2>" & formatpercent(rsdetail.fields("RecShop12m"),1) & "</td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		If IsNumeric(rsdetail.fields("CSi12m")) then 
			response.write "<td align=center><font size=2 color=red>" & formatpercent(rsdetail.fields("CSi12m"),1) & "</td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		If IsNumeric(rsdetail.fields("RecShopNum12m")) then 
			response.write "<td align=center><font size=2><B>" & formatnumber(rsdetail.fields("RecShopNum12m"),2) & "</b></td>"
		Else
			response.write "<td align=center><font size=2 color=red></td>"
		End If
		response.write "<td align=right><font size=2><b>" & rsdetail.Fields("Period12m") & "</b></td>"
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
response.write "Information compiled by CSi Complete, a national provider of CSI and call center services to business. For more information please call (800) 343-0641. Copyright 2012 CSi Complete</FONT></B></I></TD>"
response.write "</font></em>"
Response.Write "</td>"
Response.Write "</tr>"
response.write "</TABLE>"
response.write "</BODY></HTML>"
%>
