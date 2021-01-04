<% 
' 함수 정의 부분
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
더하기의 값 : <% = Add(X, Y) %> <BR>
곱하기의 값 : <% = Mul(X, Y) %> <BR>
</BODY>
</HTML>
