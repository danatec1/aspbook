<%
Set objBT = Server.CreateObject("MSWC.BrowserType")
%>
<HTML>
<HEAD>
<TITLE> BrowserType Object </TITLE>
</HEAD>
<BODY>
<% 
If objBT.frames Then %>
    �������� �������� �����Ѵ�. <BR>
<% 
Else %>
    �������� �������� �������� �ʴ´�. <BR>
<% 
End If 
If objBT.vbscript Then %>
    �������� VBScript �� �����Ѵ�.
<% 
Else %>
    �������� VBScript �� �������� �ʴ´�. <BR>
<% 
End If 
If objBT.Cookies Then %>
    �������� ��Ű�� �����Ѵ�.
<% 
Else %>
    �������� ��Ű�� �������� �ʴ´�.
<% 
End If %>
</BODY>
</HTML>
