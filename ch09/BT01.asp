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
    브라우저는 프레임을 지원한다. <BR>
<% 
Else %>
    브라우저는 프레임을 지원하지 않는다. <BR>
<% 
End If 
If objBT.vbscript Then %>
    브라우저는 VBScript 언어를 지원한다.
<% 
Else %>
    브라우저는 VBScript 언어를 지원하지 않는다. <BR>
<% 
End If 
If objBT.Cookies Then %>
    브라우저는 쿠키를 지원한다.
<% 
Else %>
    브라우저는 쿠키를 지원하지 않는다.
<% 
End If %>
</BODY>
</HTML>
