<%
' 새로운 입금액을 읽는다.
strMoney = Request.Form("varMoney")
intMoney = CInt(strMoney)
%>
<HTML>
<HEAD> 
<TITLE> Example : Blackhole Bank (Result) </TITLE>
</HEAD>
<BODY>
<I>계좌 정보</I> <BR>
새로운 입금 : <% = intMoney %> 원 <BR>
기존의 잔액 : <% = intAmount %> 원 <BR>
<BR>
<%
' 새로운 잔액을 계산한다.
'   잔액 = 입금액 + 잔액
intAmount = intMoney + intAmount
%>
잔액 : <% = intAmount %> 원 <BR>
</BODY>
</HTML>
