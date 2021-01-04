<HTML>
<HEAD> 
<TITLE> Example : HTTP Message(Result) </TITLE>
</HEAD>
<BODY>
<I>HTTP 메시지 형식</I> <BR>
<%
' HTTP Header 부분을 출력한다.
For Each Item in Request.ServerVariables("ALL_HTTP")
%>
    <% = Item & "=" %>
<%  Request.ServerVariables(Item)
%>  <BR>
<% 
Next
%>
<!-- 1 개의 공백 라인을 둔다. -->
<BR>
<%
' HTTP 데이터 부분을 출력한다.
totCount = Request.TotalBytes
binData = Request.BinaryRead(totCount)
Response.BinaryWrite binData
%>
</BODY>
</HTML>
