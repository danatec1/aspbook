<HTML>
<HEAD> 
<TITLE> Encoded Raw Data </TITLE>
</HEAD>
<BODY>
<I>Encoded Raw Data</I> <BR>
<%
' HTTP Header �κ��� ����Ѵ�.
For Each Item in Request.ServerVariables("ALL_HTTP") %>
    <% = Item & "=" %>
<%  Request.ServerVariables(Item) %>  <BR>
<% 
Next %>
<!-- 1 ���� ���� ������ �д�. -->
<BR>
<%
' HTTP ������ �κ��� ����Ѵ�.
totCount = Request.TotalBytes
binData = Request.BinaryRead(totCount)
Response.BinaryWrite binData
%>
</BODY>
</HTML>
