<HTML>
<HEAD> 
<TITLE> Example : HTTP Message(Result) </TITLE>
</HEAD>
<BODY>
<I>HTTP �޽��� ����</I> <BR>
<%
' HTTP Header �κ��� ����Ѵ�.
For Each Item in Request.ServerVariables("ALL_HTTP")
%>
    <% = Item & "=" %>
<%  Request.ServerVariables(Item)
%>  <BR>
<% 
Next
%>
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
