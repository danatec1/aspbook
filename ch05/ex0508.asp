<% 
' �����ƾ ���� �κ�
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
����� ȣ�� : X�� �� = <% Call Println(X) %>
�Ͻ��� ȣ�� : Y�� �� = <% Println(Y) %>
</BODY>
</HTML>
