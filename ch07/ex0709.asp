<%
' Hidden Field에서 읽는다.
strCard = Request.Form("card")
' 지금 입력된 데이터들을 읽는다.
strEmail = Request.Form("email")
strContent = Request.Form("content")
%>
<HTML>
<HEAD> 
<TITLE> Card Sending : Card Sending </TITLE>
</HEAD>
<BODY>
<I>카드 보내기</I> <BR>
다음과 같은 데이터를 처리하여 전자메일로 보낸다.(3rd Step) <BR>
(실제로 전자메일을 보내는 방법은 부록을 참조)  
<TABLE BORDER=1>
<TR>
    <TD> 그림 </TD>
    <TD> <% = strCard %> </TD>
</TR> <TR>
    <TD> 주소 </TD>
    <TD> <% = strEMail %> </TD>
</TR> <TR>
    <TD> 내용 </TD>
    <TD> <% = strContent %> </TD>
</TR>
</BODY>
</HTML>
