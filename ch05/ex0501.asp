<%
Jumsu = 80
%>
<HTML>
<HEAD>
<TITLE> Example : If Statements </TITLE>
</HEAD>
<BODY>
If 문들의 사용법
<HR>
<% If Jumsu >= 60 Then %>
    축하합니다 !!! 합격을 하였군요. <BR>
<% End If %>
<% If Jumsu >= 90 Then %>
    당신의 성적은 A 등급입니다.
<% ElseIf Jumsu >=80 Then %>
    당신의 성적은 B 등급입니다.
<% ElseIf Jumsu >=70 Then %>
    당신의 성적은 C 등급입니다.
<% ElseIf Jumsu >=60 Then %>
    당신의 성적은 D 등급입니다.
<% Else %>
    불합격하였습니다. 다음 기회를 이용해 주세요.
<% End IF %>   
</BODY>
</HTML>
