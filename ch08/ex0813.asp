<%
' 이 라인들은 테스트용으로 삽입되었으니 실제로 사용할 때는 삭제
Session.Timeout  = 1
Response.Expires = 0
%>
<HTML>
<HEAD> 
<TITLE> Example : Counter Test </TITLE>
</HEAD>
<BODY>
<CENTER>
당신은 <% = Application("appCounter") %>번째 방문자입니다. <BR>
<A HREF="ex0813.asp">다시읽기</A>
</CENTER>
</BODY>
</HTML>
