<% 
' ����� ���۸��ϵ��� �Ѵ�. 
' Response.Buffer = True
%>
<HTML>
<HEAD>
<TITLE> Example : Response Buffering </TITLE>
</HEAD>
<BODY>
<I>���۸�</I> <BR>
1. ���� ��ħ�� ����� ���Ҵ�. <BR>
<%  
Response.ContentType = "text/html"
Response.Flush
%>
2. ���� ������ ���ִ� ������ �Ծ���. <BR>
<%
Response.Clear
%>
3. ���� ������ �� ģ���� ������. 
</BODY>
</HTML>
<%
Response.End
%>
