<%@ Language=VBScript %>
<HTML>
<HEAD>
<SCRIPT LANGUAGE="JavaScript">
function openZoom2(url, name) {
  popupWin = window.open(url, name,
"scrollbars=yes,resizable=no,location=no,toolbar=no,top=6,left=155,width=600,height=525");
}
<!--
if (window.focus) {
  self.focus();
}
//-->
</SCRIPT>
</HEAD>
<BODY>
	<font face=arial color=brown size=2><strong>
<%
set cnList=server.CreateObject("adodb.connection")
cnlist.Open session("ConString")
'setting this for the correct survey to be displayed
	select case session("surveytypeid")
		case 1
			Response.Write "<a href=""javascript: openZoom2('http://www.csicomplete.com/reports_collision/collision_survey.htm','SURVEY');""><img src =""http://www.csicomplete.com/images/surimg2.jpg"" border =0><font style=""FONT-FAMILY: sans-serif"" fontsize=2><b>SURVEY</b></font></a>"
		case 2
			Response.Write "<a href=""javascript: openZoom2('http://www.csicomplete.com/reports_collision/collision_surveycaliber.htm','SURVEY');""><img src =""http://www.csicomplete.com/images/surimg2.jpg"" border =0><font style=""FONT-FAMILY: sans-serif"" fontsize=2><b>SURVEY</b></font></a>"
		case 3 
			Response.Write "<a href=""javascript: openZoom2('http://www.csicomplete.com/reports_collision/survey_Toyota.htm','SURVEY');""><img src =""http://www.csicomplete.com/images/surimg2.jpg"" border =0><font style=""FONT-FAMILY: sans-serif"" fontsize=2><b>SURVEY</b></font></a>"
		case 4
			Response.Write "<a href=""javascript: openZoom2('http://www.csicomplete.com/reports_collision/survey_ToyotaSpecial.htm','SURVEY');""><img src =""http://www.csicomplete.com/images/surimg2.jpg"" border =0><font style=""FONT-FAMILY: sans-serif"" fontsize=2><b>SURVEY</b></font></a>"
		case 5
			Response.Write "<a href=""javascript: openZoom2('http://www.csicomplete.com/reports_collision/collision_survey.htm','SURVEY');""><img src =""http://www.csicomplete.com/images/surimg2.jpg"" border =0><font style=""FONT-FAMILY: sans-serif"" fontsize=2><b>SURVEY</b></font></a>"
		case else
			Response.Write "<a href=""javascript: openZoom2('http://www.csicomplete.com/reports_collision/collision_survey.htm','SURVEY');""><img src =""http://www.csicomplete.com/images/surimg2.jpg"" border =0><font style=""FONT-FAMILY: sans-serif"" fontsize=2><b>SURVEY</b></font></a>"
	end select
	Response.Write "<br>"
	Response.Write "<br>"
	Response.Write "<font size=3>"	
	Response.Write session("sessionname")
	Response.Write "<br>"
%>
	</STRONG></font>
	<FONT face=Arial color=black size=3><strong>
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
		end if

	case "S"
		'RO LIST REPORTS
		%>
			</STRONG></font>
			<br>
			<font face=arial color=brown size=2><strong>
		<%
			Response.write "RO List"
			Response.Write "<br>"
		%>
			<FONT face=Arial color=black size=1.5><strong>
		<%
			strReport = "select ReportName, ReportMonth from tblCollisionReports where (CollisionID=" & session("id") & ") and (active <>'0') and (CollisionReportID=1) order by ReportMonth desc"
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
						Response.Write "<a href=/reports_collision/ROList.asp?shopid=" & session("shopid") & "&month=" & rsreport.Fields("ReportMonth") & " target=""SCREEN"">" & rsreport.Fields("reportname") & "</a>"
				end select
				Response.Write "<br>"	
				rsreport.MoveNext
			loop
			Response.Write "<br>"
			rsreport.Close

		'TREND REPORTS
		%>
			</STRONG></font>
			<font face=arial color=brown size=2><strong>
		<%
			Response.write "Trend Report"
			Response.Write "<br>"
		%>
			<FONT face=Arial color=black size=1.5><strong>
		<%
			strReport = "select ReportName, ReportMonth from tblCollisionReports where (CollisionID=" & session("id") & ") and (active <>'0') and (CollisionReportID=2) order by ReportMonth desc"
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
						Response.Write "<a href=/reports_collision/TrendCal.asp?shopid=" & session("shopid") & " target=""SCREEN"">Trend Report</a>"
					case 3
						Response.Write "<a href=/reports_collision/TrendToy.asp?shopid=" & session("shopid") & "&month=" & rsreport.Fields("ReportMonth") & " target=""SCREEN"">" & rsreport.Fields("reportname") & "</a>"
					case else
						Response.Write "<a href=/reports_collision/Trend.asp?shopid=" & session("shopid") & " target=""SCREEN"">Trend Report</a>"
				end select
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
</BODY>
</HTML>