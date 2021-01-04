<%
intX = 20
intY = 10
%>
<HTML>
<HEAD>
<TITLE> Example : On Error Statement </TITLE>
</HEAD>
<BODY>
On Error 문의 사용법
<HR>
<% 
' On Error Resume Next
P = intX / intY
P = 100 / (P - 2)  
R = intX mod intY 
%>
P의 값은 : <% = P %> <BR>
R의 값은 : <% = R %>
</BODY>
</HTML>
