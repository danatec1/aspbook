<%
strMoney = Request.Form("varMoney")
' 쿠키값을 설정한다.(클라이언트로 전송한다.)
Response.Cookies("cooAmount") = strMoney
%>
<HTML>
<HEAD> 
<TITLE> Cookie Bank </TITLE>
</HEAD>
<BODY>
<I>계좌 정보</I> <BR>
잔액 : <% = strMoney %> 원 
<FORM METHOD="post" ACTION="ex0712.asp">
금액 : <INPUT TYPE="text" NAME="varMoney" SIZE=10> 원 <BR>
<INPUT TYPE="submit" VALUE="입금">
</FORM>
</BODY>
</HTML>
