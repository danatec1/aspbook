<% 
' �Լ� ���� �κ�
Function Add(v1, v2)
    v3 = v1 + v2
    Add = v3
End Function 
Function Mul(v1, v2)
    v3 = v1 * v2
    Mul = v3
End Function 
%>
<HTML>
<HEAD>
<TITLE> Example : Function </TITLE>
</HEAD>
<BODY BGCOLOR=white>
<%
X = 10
Y = 20
%>
���ϱ��� �� : <% = Add(X, Y) %> <BR>
���ϱ��� �� : <% = Mul(X, Y) %> <BR>
</BODY>
</HTML>
