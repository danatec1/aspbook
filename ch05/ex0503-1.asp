<%
intX = 20
intY = 10
%>
<HTML>
<HEAD>
<TITLE> Example : On Error Statement </TITLE>
</HEAD>
<BODY>
On Error ���� ����
<HR>
<% 
' On Error Resume Next
P = intX / intY
P = 100 / (P - 2)  
R = intX mod intY 
%>
P�� ���� : <% = P %> <BR>
R�� ���� : <% = R %>
</BODY>
</HTML>
