<%@ Language=VBScript %>
<% if session("id")=""  then%>

<html>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>CSiComplete Logon</TITLE>
<script Language="VBScript"> 
Function document_onkeypress 
if window.event.keycode = 13 then 
     GO_Onclick() 
End if 
End function 
 
Sub GO_OnClick 
Document.DataForm.submit 
end sub 
</script>

</HEAD>

<body>

<form name='dataform' action='logon.htm'> 

<P>&nbsp</p>
<P><FONT face=Arial color=brick><strong>You must log on to the CSi Complete 
Reporting website to view this page.</strong></FONT></P>
<P>&nbsp</p>
<P>&nbsp</p>

<p align="center"><input type='button' name='GO' value='Go To CSiComplete Logon'> </p>
</form> 
 
</body>

</html>

<% else%>
	

<html>

<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 3.0">
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=iso-8859-1"> <title>CSi Complete Reports</title> </head>
<body>

  <table border="0" cellpadding="0" cellspacing="0" style="BORDER-COLLAPSE: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
  <tr>
    <td width="36%" align="left">
    <IMG style="WIDTH: 173px; HEIGHT: 44px" height=44 hspace=1 src="../images/csilogoSMALL.jpg" width=173 vspace=1 border=0><br>
	</td>
	
<td width="15%" align="right"><strong><font style="FONT-FAMILY: sans-serif" fontsize=2>
	&nbsp;<A href="/reports_collision/display.asp" target="screen">Main Menu</A> </font></strong>


<td width="15%" align="left"><strong><font  
      style="FONT-FAMILY: sans-serif" fontsize=2>&nbsp;
      <A href="http://www.csicomplete.com/reports_collision/logout.asp" target=_top>Log 
      Off</A> </font></strong>
</td>
</tr>
</table>
</body>
</html>
<% end if%>