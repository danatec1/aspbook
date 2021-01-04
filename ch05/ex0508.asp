<% 
' 서브루틴 정의 부분
Sub Println(msg)
%>
    <% = msg %> <BR>
<%
End Sub 
%>
<HTML>
<HEAD>
<TITLE> Example : Subroutine </TITLE>
</HEAD>
<BODY BGCOLOR=white>
<%
X = 100
Y = 200
%>
명시적 호출 : X의 값 = <% Call Println(X) %>
암시적 호출 : Y의 값 = <% Println(Y) %>
</BODY>
</HTML>
