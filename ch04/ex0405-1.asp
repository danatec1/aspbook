<%
Option Explicit
' 변수를 암시적으로 선언하여 사용
strX = "안녕하세요 !!!"
' 변수를 명시적으로 선언하여 사용
DIM intY
intY = 5
%>
<HTML>
<HEAD>
<TITLE> Example : Variable </TITLE>
</HEAD>
<BODY BGCOLOR=white>
스트링 변수값 : <% = strX %> <BR>
정수 변수값 : <% = intY %> <BR>
</BODY>
</HTML>

