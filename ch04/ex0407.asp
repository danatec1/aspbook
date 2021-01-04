<%
intX = 11
intY = 2 
Sum = intX + intY
Sbt =  intX - intY
Mul = intX * intY
Div = intX / intY
Epo = intX ^ intY
Quo = intX \ intY
Rmi = intX Mod intY 
%>
<HTML>
<HEAD>
<TITLE> Example : Arithmatic Operators </TITLE>
</HEAD>
<BODY BGCOLOR=white>
수치 연산자들
<HR>
덧셈의 결과 : <% = Sum %> <BR>
뺄셈의 결과 : <% = Sbt %> <BR>
곱셈의 결과 : <% = Mul %> <BR>
나눗셈의 결과 : <% = Div %> <BR>
지수 연산의 결과 : <% = Epo %> <BR>
정수 나눗셈(몫)의 결과 : <% = Quo %> <BR>
Modules 연산(나머지)의 결과 : <% = Rmi %>
</BODY>
</HTML>
