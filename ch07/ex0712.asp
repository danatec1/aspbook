<%
' 쿠키값을 읽는다. (자동으로 서버로 전송된다.)
strOldAmount = Request.Cookies("cooAmount")
intOldAmount = CInt(strOldAmount)
' 새로운 입금액을 읽는다.
strMoney = Request.Form("varMoney")
intMoney = CInt(strMoney)
' 새로운 잔액을 계산한다.
'   잔액 = 입금액 + 기존의 잔액
intAmount = intMoney + intOldAmount
' 쿠키값을 재설정한다.(클라이언트의 쿠키값을 변경한다.)
Response.Cookies("cooAmount") = intAmount
%>
<HTML>
<HEAD> 
<TITLE> Cookie Bank </TITLE>
</HEAD>
<BODY>
<I>계좌 정보</I> <BR>
새로운 입금 : <% = intMoney %> 원 <BR>
기존의 잔액 : <% = intOldAmount %> 원 <BR>
잔액 : <% = intAmount %> 원 <BR>
<FORM METHOD="post" ACTION="ex0712.asp">
금액 : <INPUT TYPE="text" NAME="varMoney" SIZE=10> 원 <BR>
<INPUT TYPE="submit" VALUE="추가입금"> 
</FORM>
</BODY>
</HTML>
