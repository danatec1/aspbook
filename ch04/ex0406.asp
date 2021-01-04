<%
intX = 100
dblY = 1.2345 
strS = "¾È³çÇÏ¼¼¿ä !!!"
dtmD = Date()
%>
<HTML>
<HEAD>
<TITLE> Example : Variable Sub Type</TITLE>
</HEAD>
<BODY BGCOLOR=white>
intX : <% = TypeName(intX) %> : <% = intX %> <BR> 
dblY : <% = TypeName(dblY) %> : <% = dblY %> <BR>
strS : <% = TypeName(strS) %> : <% = strS %> <BR>
dtmD : <% = TypeName(dtmD) %> : <% = dtmD %> <BR>
<BR>
<% strX = CStr(intX) %>
strX : <% = TypeName(strX) %> : <% = strX %> <BR> 
strY : <% = TypeName(strY) %> : <% = strY %> <BR> 
</BODY>
</HTML>

