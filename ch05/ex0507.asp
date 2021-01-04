<HTML>
<HEAD>
<TITLE> Example : While...Wend Statement </TITLE>
</HEAD>
<BODY BGCOLOR=white>
1 ~ 10까지 합하는 과정 <BR>
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
결과(총합) : <% = intSum %>
</BODY>
</HTML>
