<HTML>
<HEAD>
<TITLE> Example : While...Wend Statement </TITLE>
</HEAD>
<BODY BGCOLOR=white>
1 ~ 10���� ���ϴ� ���� <BR>
<% 
intSum = 0
intN = 1
While (intN < 11)  
    intSum = intSum + intN
%>
    <%= intSum %> &nbsp;  
<% 
    intN = intN + 1
Wend 
%>
<BR>
���(����) : <% = intSum %>
</BODY>
</HTML>
