<HTML>
<HEAD>
<TITLE> Example : Do...Loop Statement </TITLE>
</HEAD>
<BODY BGCOLOR=white>
1 ~ 10까지 모든 숫자를 출력하기 <BR>
<% intCount = 1 
    Do Until intCount > 10 %>
        <% = intCount %> &nbsp;
<%      intCount = intCount + 1 
    Loop %>
<BR>
10 ~ 1까지 짝수만 출력하기 <BR>
<% intCount = 10
    Do %>
	<% = intCount %> &nbsp;  
<%      intCount = intCount - 2
        If intCount < 2 Then
	    Exit Do
        End If
    Loop Until False %>
</BODY>
</HTML>
