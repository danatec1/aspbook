<!-- #include file="SBConn.inc" -->
<%
' ���α׷��� �μ��� �д´�.
strMID = Request.QueryString("MID")
intMID = CInt(strMID)
' �����ͺ��̽����� �ش� �Խù�(���ڵ�)�� �д´�. 
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
sqlQuery = "SELECT * FROM Message WHERE MID=" & intMID
Set objRS = objConn.Execute(sqlQuery)
' �Խù��� �ʵ� ������ �д´�. 
strSubject = objRS("MSubject")
' �Խ��� ������ CR ���ڸ� "<BR>" �� �ٲ۴�. 
strContent = Replace(objRS("MContent"), chr(13), "<BR>")
strUName = objRS("MUName")
%> 
<HTML>
<HEAD> 
<TITLE> Message View </TITLE>
</HEAD>
<BODY>
<I>���б�</I><BR>
<TABLE border=1>
<TR>
    <TD> �̸� </TD>
    <TD> <%= strUName %> </TD>	
</TR> <TR>
    <TD> ���� </TD>
    <TD> <%= strSubject %> </TD>
</TR> <TR>
    <TD> ���� </TD>
    <TD> <%= strContent %> </TD>	
</TR>
</TABLE>
<A HREF="SBList.asp">��Ϻ���</A> 
<A HREF="SBUForm.asp?MID=<%= strMID %>">�����ϱ�</A> 
<A HREF="SBDelete.asp?MID=<%= strMID %>">�����ϱ�</A> 
</BODY>
</HTML>
<%
objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
%>
