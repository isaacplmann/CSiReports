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
<BODY><font face =arial color=brown size=4><strong>
<%
set cnList=server.CreateObject("adodb.connection")
cnlist.Open session("ConString")
set rsROList = server.CreateObject("adodb.recordset")

'get the shopid to grab the session shopid and session name
strROList = "select ShopID, Name, CollisionID, ParentID from tblCollisionLogons where collisionid=" & session("id")
rsrolist.Open strrolist, cnlist
session("ShopID") = rsrolist.Fields("ShopID")
rsrolist.Close

	Response.Write session("sessionname")
	Response.Write "<br>"
	Response.Write "<br>"
%>
	</STRONG></font>
	<font face=arial color=green size=2><strong>
<%
	if session("sessionparentid") = 8 then
		Response.Write "<a href=""javascript: openZoom2('http://www.csicomplete.com/reports_collision/collision_surveycaliber.htm','SURVEY');""><img src =""http://www.csicomplete.com/images/surimg2.jpg"" border =0><font style=""FONT-FAMILY: sans-serif"" fontsize=2><b>SURVEY</b></font></a>"
	else
		Response.Write "<a href=""javascript: openZoom2('http://www.csicomplete.com/reports_collision/collision_survey.htm','SURVEY');""><img src =""http://www.csicomplete.com/images/surimg2.jpg"" border =0><font style=""FONT-FAMILY: sans-serif"" fontsize=2><b>SURVEY</b></font></a>"
	end if

	'Response.Write "<a href=""javascript: openZoom2('http://www.csicomplete.com/reports_collision/collision_survey.htm','SURVEY');""><img src =""http://www.csicomplete.com/images/surimg2.jpg"" border =0><font style=""FONT-FAMILY: sans-serif"" fontsize=2><b>SURVEY</b></font></a>"
	Response.Write "<br>"
	Response.Write "<br>"
%>
	</STRONG></font>
	<font face=arial color=brown size=2><strong>
<%
	Response.write "RO List"
	Response.Write "<br>"
%>
	<FONT face=Arial color=black size=1.5><strong>
<%

	strReport = "select ReportID, ReportName, ReportMonth, ReportASP from tblCollisionReports where CollisionID=" & session("id") & " and active <>'0' order by ReportMonth desc"
	set rsReport = server.CreateObject("adodb.recordset")
	rsreport.Open strreport,cnlist
	
	if rsreport.EOF then
		Response.Write "There are no reports to display."
		Response.Write "<br>"
	end if

	do until rsreport.EOF
		'get the report for that month
	if session("sessionparentid")=8 then
		Response.Write "<a href=/reports_collision/ROListCaliber.asp?shopid=" & session("shopid") & "&month=" & rsreport.Fields("ReportMonth") & " target=""SCREEN"">" & rsreport.Fields("reportname") & "</a>"
	else
		Response.Write "<a href=/reports_collision/ROList.asp?shopid=" & session("shopid") & "&month=" & rsreport.Fields("ReportMonth") & " target=""SCREEN"">" & rsreport.Fields("reportname") & "</a>"
	end if
		Response.Write "<br>"
			
		rsreport.MoveNext
	loop
	Response.Write "<br>"
	rsreport.Close
	
%>
</BODY>
</HTML>