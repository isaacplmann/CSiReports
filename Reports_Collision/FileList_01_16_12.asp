<%@ Language=VBScript %>
<html>
<head>
<script language="JavaScript">
function openZoom2(url, name) {
  popupWin = window.open(url, name,
"scrollbars=yes,resizable=no,location=no,toolbar=no,top=6,left=155,width=600,height=525");
}
<!--
if (window.focus) {
  self.focus();
}
//-->
</script>
</head>
<body>
	<font face=arial color=brown size=2><strong>
<%
set cnList=server.CreateObject("adodb.connection")
cnlist.Open session("ConString")
'get the id of the shop
ID = session("ID")

	if id = 252 then 'use autonation logo
		response.Write "<IMG alt="""" src=../images/Autonation_Logo.jpg>"
		Response.Write "<br>"
		Response.Write "<br>"
	else 'otherwise, just use name
		Response.Write "<font size=3>"	
		Response.Write session("sessionname")
		Response.Write "<br>"
		Response.Write "<br>"
	end if


	'if its autionation, show wav files
	if id = 252 then 
		Response.Write "<a href=http://www.csicomplete.com/reports_collision/Autonation_WavFiles.asp target=""SCREEN""><img src =""http://www.csicomplete.com/images/sound.gif"" border =0><font style=""FONT-FAMILY: sans-serif"" fontsize=2><b>Hot Sheet Summary</b></font></a>"
		Response.Write "<br>"
		Response.Write "<br>"
	end if
	
'setting this for the correct survey to be displayed
ParentID = session("sessionparentid")
MSO = session("sessionmso")
MSONew = session("sessionmsonew")
MSOIns = session("sessionmsoins")
Name = session("sessionname")

If (MSOIns=-1) then 'use mso ins survey
	Response.Write "<a href=""javascript: openZoom2('http://www.csicomplete.com/reports_collision/survey_mso_ins.htm','Survey');""><img src =""http://www.csicomplete.com/images/surimg2.jpg"" border =0><font style=""FONT-FAMILY: sans-serif"" fontsize=2><b>Survey</b></font></a>" 
Else
	If (MSO=-1) then 'use mso survey
            If (ParentID=77) or (Name="Gerber")  then
			 	Response.Write "<a href=""javascript: openZoom2('http://www.csicomplete.com/reports_collision/survey_gerber.htm','Survey');""><img src =""http://www.csicomplete.com/images/surimg2.jpg"" border =0><font style=""FONT-FAMILY: sans-serif"" fontsize=2><b>Survey</b></font></a>" 
			Else
				If (ParentID=8) or (Name="calibercollision") then
					Response.Write "<a href=""javascript: openZoom2('http://www.csicomplete.com/reports_collision/survey_mso_new.htm','Survey');""><img src =""http://www.csicomplete.com/images/surimg2.jpg"" border =0><font style=""FONT-FAMILY: sans-serif"" fontsize=2><b>Survey</b></font></a>" 
				Else
					Response.Write "<a href=""javascript: openZoom2('http://www.csicomplete.com/reports_collision/survey_mso.htm','Survey');""><img src =""http://www.csicomplete.com/images/surimg2.jpg"" border =0><font style=""FONT-FAMILY: sans-serif"" fontsize=2><b>Survey</b></font></a>" 
				End If
			End If				
    Else
    select case session("surveytypeid")
        case 1
            Response.Write "<a href=""javascript: openZoom2('http://www.csicomplete.com/reports_collision/collision_survey.htm','Survey');""><img src =""http://www.csicomplete.com/images/surimg2.jpg"" border =0><font style=""FONT-FAMILY: sans-serif"" fontsize=2><b>Survey</b></font></a>"
        case 2
	        Response.Write "<a href=""javascript: openZoom2('http://www.csicomplete.com/reports_collision/collision_surveycaliber.htm','Survey');""><img src =""http://www.csicomplete.com/images/surimg2.jpg"" border =0><font style=""FONT-FAMILY: sans-serif"" fontsize=2><b>Survey</b></font></a>"
        case 3 
	        Response.Write "<a href=""javascript: openZoom2('http://www.csicomplete.com/reports_collision/survey_Toyota.htm','Survey');""><img src =""http://www.csicomplete.com/images/surimg2.jpg"" border =0><font style=""FONT-FAMILY: sans-serif"" fontsize=2><b>Survey</b></font></a>"
        case 4
	        Response.Write "<a href=""javascript: openZoom2('http://www.csicomplete.com/reports_collision/survey_ToyotaSpecial.htm','Survey');""><img src =""http://www.csicomplete.com/images/surimg2.jpg"" border =0><font style=""FONT-FAMILY: sans-serif"" fontsize=2><b>Survey</b></font></a>"
        case 5
	        Response.Write "<a href=""javascript: openZoom2('http://www.csicomplete.com/reports_collision/collision_survey.htm','Survey');""><img src =""http://www.csicomplete.com/images/surimg2.jpg"" border =0><font style=""FONT-FAMILY: sans-serif"" fontsize=2><b>Survey</b></font></a>"
		case 7
	        If (ParentID=77) or (Name="Gerber") then
				Response.Write "<a href=""javascript: openZoom2('http://www.csicomplete.com/reports_collision/survey_gerber.htm','Survey');""><img src =""http://www.csicomplete.com/images/surimg2.jpg"" border =0><font style=""FONT-FAMILY: sans-serif"" fontsize=2><b>Survey</b></font></a>" 
			Else
				If (ParentID=8) or (Name="calibercollision") then
					Response.Write "<a href=""javascript: openZoom2('http://www.csicomplete.com/reports_collision/survey_mso_new.htm','Survey');""><img src =""http://www.csicomplete.com/images/surimg2.jpg"" border =0><font style=""FONT-FAMILY: sans-serif"" fontsize=2><b>Survey</b></font></a>" 
				Else
					Response.Write "<a href=""javascript: openZoom2('http://www.csicomplete.com/reports_collision/survey_mso.htm','Survey');""><img src =""http://www.csicomplete.com/images/surimg2.jpg"" border =0><font style=""FONT-FAMILY: sans-serif"" fontsize=2><b>Survey</b></font></a>" 
				End If
				
			End If
        case else
	        Response.Write "<a href=""javascript: openZoom2('http://www.csicomplete.com/reports_collision/collision_survey.htm','Survey');""><img src =""http://www.csicomplete.com/images/surimg2.jpg"" border =0><font style=""FONT-FAMILY: sans-serif"" fontsize=2><b>Survey</b></font></a>"
    end select
    
    
    end if
end if

	Response.Write "<br>"
	Response.Write "<br>"
%>
	</strong></font>
	<font face=Arial color=black size=3><strong>
<%

set rsPReports = server.CreateObject("adodb.recordset")
set rsPRegions = server.CreateObject("adodb.recordset")
set rsRReports = server.CreateObject("adodb.recordset")
set rsRShops = server.CreateObject("adodb.recordset")
set rsSReports = server.CreateObject("adodb.recordset")

select case session("clienttype")
	case "P"
		if session("clienttype")="P" then
		'list the parent reports if there are any
			Response.Write "<font size=1>"	
			strPReport = "select ReportName, ReportMonth, ReportASP from tblCollisionReports where (CollisionID=" & session("id") & ") and (active <>'0') order by CollisionReportID, ReportMonth desc"
			rsPReports.Open strPReport,cnlist
			do until rsPReports.EOF
				Response.Write "<a href=/reports_collision/" & rspreports.Fields("ReportASP") & "?Month=" & rspreports.Fields("ReportMonth") & " target=""SCREEN"">" & rsPReports.Fields("reportname") & "</a><br>"		
				rspreports.MoveNext
			loop
			rsPReports.close

			Response.Write "</strong>"
			Response.Write "<br>"
			Response.Write "<br>"
		
			'are there regions?
			strPRegions = "select Name, CollisionID from tblCollisionLogons where (Active=-1) and (ParentID=" & session("id") & ") and (Type='R') order by Name"
			rsPRegions.open strPRegions, cnlist
			if rspregions.EOF then
				'there are no regions, look for shops
				strRShops = "select Name, CollisionID, ShopID from tblCollisionLogons where (Active=-1) and (ParentID=" & session("id") & ") and (Type='S') order by Name"
				rsRShops.open strrshops, cnlist
					do until rsrshops.EOF
						Response.Write "<br>"
						Response.Write "<font size=2 color=black><b>" & rsRShops.fields("Name") & "</b><br>"
						Response.Write "<font size=1>"
						'get the reports for the shop
						strSReports = "select ReportName, ReportMonth, ReportASP from tblCollisionReports where (CollisionID=" & rsRShops.fields("CollisionID") & ") and (active <>'0') order by CollisionReportID, ReportMonth desc"		
						rsSReports.open strSReports, cnlist
						do until rssreports.EOF
							Response.Write "<a href=/reports_collision/" & rssreports.Fields("ReportASP") & "?ShopID=" & rsrshops.Fields("ShopID") & "&Month=" & rssreports.Fields("ReportMonth") & " target=""SCREEN"">" & rssreports.Fields("reportname") & "</a><br>"		
							rssreports.MoveNext
						loop
						rssreports.Close
						rsrshops.MoveNext
					loop
					rsrshops.Close
			else
				do until rsPRegions.eof
					Response.Write "<font color=brown size=2><strong>"
					Response.Write rspregions.fields("Name") & "<br>"
					
					'grab any region reports
					Response.Write "</strong><font size=1>"
					strRReports = "select ReportName, ReportMonth, ReportASP from tblCollisionReports where (CollisionID=" & rspregions.fields("CollisionID") & ") and (active <>'0') order by CollisionReportID, ReportMonth desc"		
					rsRReports.open strRReports, cnlist
					do until rsRReports.EOF
						Response.Write "<a href=/reports_collision/" & rsrreports.Fields("ReportASP") & "?Month=" & rsrreports.Fields("ReportMonth") & " target=""SCREEN"">" & rsrreports.Fields("reportname") & "</a><br>"		
						rsrreports.MoveNext
					loop
					rsrreports.Close
					Response.Write "</strong>"
					
					'grab the shops in each region
					strRShops = "select Name, CollisionID, ShopID from tblCollisionLogons where (Active=-1) and (ParentID=" & session("id") & ") and (RegionID=" & rsPRegions.Fields("CollisionID") & ") and (Type='S') order by Name"
					rsRShops.open strrshops, cnlist
					do until rsrshops.EOF
						Response.Write "<br>"
						Response.Write "<font size=2 color=black>" & rsRShops.fields("Name") & "<br>"
						Response.Write "<font size=1>"
						'get the reports for the shop
						strSReports = "select ReportName, ReportMonth, ReportASP from tblCollisionReports where (CollisionID=" & rsRShops.fields("CollisionID") & ") and (active <>'0') order by CollisionReportID, ReportMonth desc"		
						rsSReports.open strSReports, cnlist
						do until rssreports.EOF
							Response.Write "<a href=/reports_collision/" & rssreports.Fields("ReportASP") & "?ShopID=" & rsrshops.Fields("ShopID") & "&Month=" & rssreports.Fields("ReportMonth") & " target=""SCREEN"">" & rssreports.Fields("reportname") & "</a><br>"		
							rssreports.MoveNext
						loop
						rssreports.Close
						rsrshops.MoveNext
					loop
					rsrshops.Close
					rsPRegions.movenext
					Response.Write "<br>"
				loop
			end if
			rspregions.Close
		End If
	case "S"
		'RO LIST REPORTS
		%>
			</strong></font>
			<br>
			<font face=arial color=brown size=2><strong>
		<%
			Response.write "RO List"
			Response.Write "<br>"
		%>
			<font face=Arial color=black size=1.5><strong>
		<%
			strReport = "select ReportName, ReportMonth from tblCollisionReports where (CollisionID=" & session("id") & ") and (active <>'0') and (CollisionReportID=1 or CollisionReportID=15 or CollisionReportID=27 or CollisionReportID=33) order by ReportMonth desc"
			'strReport = "select ReportName, ReportMonth from tblCollisionReports where (CollisionID=" & session("id") & ") and (active <>'0') order by ReportMonth desc"
			set rsReport = server.CreateObject("adodb.recordset")
			rsreport.Open strreport,cnlist
			
			if rsreport.EOF then
				Response.Write "There are no reports to display."
				Response.Write "<br>"
			end if
			
			

			do until rsreport.EOF
				'get the report for that month
			    select case session("surveytypeid")
				    case 2
					    Response.Write "<a href=/reports_collision/ROListCaliber.asp?shopid=" & session("shopid") & "&month=" & rsreport.Fields("ReportMonth") & " target=""SCREEN"">" & rsreport.Fields("reportname") & "</a>"
				    case 3
					    Response.Write "<a href=/reports_collision/ROListToyota.asp?shopid=" & session("shopid") & "&month=" & rsreport.Fields("ReportMonth") & " target=""SCREEN"">" & rsreport.Fields("reportname") & "</a>"
				    case else
						If (MSOIns=-1) then
							Response.Write "<a href=/reports_collision/ROListIns.asp?shopid=" & session("shopid") & "&month=" & rsreport.Fields("ReportMonth") & " target=""SCREEN"">" & rsreport.Fields("reportname") & "</a>"
						Else
							If (ParentID=8) then
								Response.Write "<a href=/reports_collision/ROListNew.asp?shopid=" & session("shopid") & "&month=" & rsreport.Fields("ReportMonth") & " target=""SCREEN"">" & rsreport.Fields("reportname") & "</a>"
							Else
								Response.Write "<a href=/reports_collision/ROList.asp?shopid=" & session("shopid") & "&month=" & rsreport.Fields("ReportMonth") & " target=""SCREEN"">" & rsreport.Fields("reportname") & "</a>"
							End If
						End If
			    end select
				
				Response.Write "<br>"	
				rsreport.MoveNext
			loop
			Response.Write "<br>"
			rsreport.Close

		'TREND REPORTS
		%>
			</strong></font>
			<font face=arial color=brown size=2><strong>
		<%
			Response.write "Trend Report"
			Response.Write "<br>"
		%>
			<font face=Arial color=black size=1.5><strong>
		<%
			strReport = "select ReportName, ReportMonth from tblCollisionReports where (CollisionID=" & session("id") & ") and (active <>'0') and (CollisionReportID=2 or CollisionReportID=16 or CollisionReportID=28 or CollisionReportID=34) order by ReportMonth desc"
			set rsReport = server.CreateObject("adodb.recordset")
			rsreport.Open strreport,cnlist
			
			if rsreport.EOF then
				Response.Write "There are no reports to display."
				Response.Write "<br>"
			end if

			do until rsreport.EOF
				'get the report for that month
				ParentID = session("sessionparentid")
				'MSO = session("sessionmso")
	            'If (MSO=-1) then 'use mso survey
				'If session("MSO")=-1 then
                '         Response.Write "<a href=/reports_collision/trendmso.asp?shopid=" & session("shopid") & " target=""SCREEN"">Trend Report</a>"
                'Else
				select case session("surveytypeid")
					case 2
						Response.Write "<a href=/reports_collision/TrendCal.asp?shopid=" & session("shopid") & " target=""SCREEN"">Trend Report</a>"
					case 3
						Response.Write "<a href=/reports_collision/TrendToy.asp?shopid=" & session("shopid") & "&month=" & rsreport.Fields("ReportMonth") & " target=""SCREEN"">" & rsreport.Fields("reportname") & "</a>"
					case 7
						If (MSOIns=-1) then
							Response.Write "<a href=/reports_collision/trendmsoins.asp?shopid=" & session("shopid") & " target=""SCREEN"">Trend Report</a>"
						Else
							If ParentID=8 Then
								Response.Write "<a href=/reports_collision/trendmsonew.asp?shopid=" & session("shopid") & " target=""SCREEN"">Trend Report</a>"
							Else
								Response.Write "<a href=/reports_collision/trendmso.asp?shopid=" & session("shopid") & " target=""SCREEN"">Trend Report</a>"
							End If
						End If
					case else
						If (MSOIns=-1) then
							Response.Write "<a href=/reports_collision/trendmsoins.asp?shopid=" & session("shopid") & " target=""SCREEN"">Trend Report</a>"
						Else
							Response.Write "<a href=/reports_collision/Trend.asp?shopid=" & session("shopid") & " target=""SCREEN"">Trend Report</a>"
						End If
				end select
				'end if
				Response.Write "<br>"
				rsreport.MoveNext
			loop
			Response.Write "<br>"
			rsreport.Close
	case "R"
		if session("clienttype")="R" then
		'list the region reports if there are any
			Response.Write "<font size=1>"	
			strRReport = "select ReportName, ReportMonth, ReportASP from tblCollisionReports where (CollisionID=" & session("id") & ") and (active <>'0') order by CollisionReportID, ReportMonth desc"
			rsRReports.Open strRReport,cnlist
			do until rsRReports.EOF
				Response.Write "<a href=/reports_collision/" & rsrreports.Fields("ReportASP") & "?Month=" & rsrreports.Fields("ReportMonth") & " target=""SCREEN"">" & rsrReports.Fields("reportname") & "</a><br>"		
				rsrreports.MoveNext
			loop
			rsrReports.close

			Response.Write "</strong>"
			Response.Write "<br>"
			
			'grab all of the shops in the region
			strRShops = "select Name, CollisionID, ShopID from tblCollisionLogons where (Active=-1) and (RegionID=" & session("id") & ") and (Type='S') order by Name"
			rsRShops.open strRShops, cnlist

			do until rsRShops.eof
				Response.Write "<font color=brown size=2><strong>"
				Response.Write "<br>" & rsrshops.fields("Name") & "<br>"
				
				'grab any shop reports
				Response.Write "</strong><font size=1>"
				strSReports = "select ReportName, ReportMonth, ReportASP from tblCollisionReports where (CollisionID=" & rsrshops.fields("CollisionID") & ") and (active <>'0') order by CollisionReportID, ReportMonth desc"		
				rsSReports.open strSReports, cnlist
				do until rsSReports.EOF
					Response.Write "<a href=/reports_collision/" & rssreports.Fields("ReportASP") & "?shopID=" & rsrshops.Fields("shopid") & "&Month=" & rssreports.Fields("ReportMonth") & " target=""SCREEN"">" & rssreports.Fields("reportname") & "</a><br>"		
					rssreports.MoveNext
				loop
				rssreports.Close
				Response.Write "</strong>"
				rsrshops.MoveNext
			loop
			rsrshops.Close
		end if
end select


%>
</body>
</html>