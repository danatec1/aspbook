<%
strTable = "Users"
strName = "OGSon"
SQLQuery = "SELECT * FROM " & strTable
SQLQuery = SQLQuery & " WHERE UName=" + strName

strX = "10"
strY = "20"
intX = 10
Sum1 = strX + strY
Sum2 = intX + strY
%>
<HTML>
<HEAD>
<TITLE> Example : Concatenation Operators </TITLE>
</HEAD>
<BODY>
연결 연산자들
<HR>
권장 사용 방법 : <% = SQLQuery %> <BR>
잘못된(예상치 못한) 결과 (문자열 + 문자열) <BR> 
SUM1 : <% = Sum1 %> <BR>
정상적인 결과 (정수 + 문자열) <BR>
SUM2 : <% = Sum2 %> 
</BODY>
</HTML>
