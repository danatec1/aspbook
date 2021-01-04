<%
strMoney = Request.Form("varMoney")
intAmount = CInt(strMoney)
%>
<HTML>
<HEAD> 
<TITLE> Example : Blackhole Bank </TITLE>
</HEAD>
<BODY>
<I>계좌 정보</I> <BR>
잔액 : <% = intAmount %> 원 
<FORM METHOD="post" ACTION="ex0706.asp">
금액 : <INPUT TYPE="text" NAME="varMoney" SIZE=10> 원 <BR>
<INPUT TYPE="submit" VALUE="입금">
</FORM>
</BODY>
</HTML>
