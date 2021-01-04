<HTML>
<HEAD> 
<TITLE> Example : Do...Loop Statement </TITLE>
</HEAD>
<BODY>
10 ~ 1까지 모든 숫자를 출력하기 <BR>
<% intCount = 10 
    Do While intCount >= 1 %>
        <% = intCount %> &nbsp;
<%      intCount = intCount - 1 
    Loop %>
<BR>
1 ~ 10 까지 짝수만 출력하기 <BR>
<% intCount = 2
    Do %>
	<% = intCount %> &nbsp;  
<%      intCount = intCount + 2
    Loop While (intCount <= 10) %>
</BODY>
</HTML>
