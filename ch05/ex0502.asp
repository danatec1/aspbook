<%
intN = 7
%>
<HTML>
<HEAD>
<TITLE> Example : Select Statement </TITLE>
</HEAD>
<BODY>
Select 문의 사용법
<HR>
<% 
Select Case intN 
    Case 1, 3, 5, 7, 9 
%>
    홀수입니다. 
<% Case 0, 2, 4, 6, 8 %>
    짝수입니다. 
<% Case Else %>
    입력 데이터가 한자리수가 아닙니다. 
<% 
End Select 
%>
</BODY>
</HTML>
