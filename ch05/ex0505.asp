<HTML>
<HEAD> 
<TITLE> Example : Do...Loop Statement </TITLE>
</HEAD>
<BODY>
10 ~ 1���� ��� ���ڸ� ����ϱ� <BR>
<% intCount = 10 
    Do While intCount >= 1 %>
        <% = intCount %> &nbsp;
<%      intCount = intCount - 1 
    Loop %>
<BR>
1 ~ 10 ���� ¦���� ����ϱ� <BR>
<% intCount = 2
    Do %>
	<% = intCount %> &nbsp;  
<%      intCount = intCount + 2
    Loop While (intCount <= 10) %>
</BODY>
</HTML>
