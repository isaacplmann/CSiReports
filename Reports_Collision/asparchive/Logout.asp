<%@ Language=VBScript %>
<%
session("id")="-1"
session("sessionname") = "Invalid Logon"
session("sessionparentid") = "-1"
session("sessionparentname") = "Press LogOff to return to log in"
Response.Redirect("Logon.htm")
%><HTML><HEAD><META content="MSHTML 6.00.2713.1100" name=GENERATOR></HEAD>
<BODY>LOGON VERIFY PAGE</BODY></HTML>