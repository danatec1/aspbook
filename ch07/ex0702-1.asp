<% 
' 출력을 버퍼링하도록 한다. 
' Response.Buffer = True
%>
<HTML>
<HEAD>
<TITLE> Example : Response Buffering </TITLE>
</HEAD>
<BODY>
<I>버퍼링</I> <BR>
1. 오늘 아침은 기분이 좋았다. <BR>
<%  
Response.ContentType = "text/html"
Response.Flush
%>
2. 오늘 점심은 맛있는 음식을 먹었다. <BR>
<%
Response.Clear
%>
3. 오늘 저녁은 옛 친구를 만났다. 
</BODY>
</HTML>
<%
Response.End
%>
