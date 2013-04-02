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

<form name='dataform' action='..\reports\logon.htm'> 

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
<base target="SCREEN">

<frameset frameborder="0" rows="50,*">
  <frame frameborder="0" MARGINHEIGHT="0" MARGINWIDTH="0" NAME="BANNER" SCROLLING="no"
  SRC="header.html" height=600>
  <frameset frameborder="0" cols="175,*">
    <frame ALIGN="LEFT" FRAMEBORDER="1" MARGINHEIGHT="0" MARGINWIDTH="0" NAME="MENU"
    SCROLLING="as needed" SRC="FileList.asp">
    <frame ALIGN="LEFT" FRAMEBORDER="1" MARGINHEIGHT="0" MARGINWIDTH="0" NAME="SCREEN"
    SCROLLING="as needed" SRC="display.asp">
  </frameset>
  <noframes>
  <body>
  </body>
  </noframes>
</frameset>
</html>
<% end if%>