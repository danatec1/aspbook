<!-- #include file="SJConf.inc" -->
<%
' ���α׷��� �μ��� �д´�.
strMID = Request.QueryString("MID")
intMID = CInt(strMID)
' �����ͺ��̽����� �ش� �ڷ�(���ڵ�)�� �д´�. 
set objConn = Server.CreateObject("ADODB.Connection")
objConn.open ConnString
sqlQuery = "SELECT * FROM JMessage WHERE MID=" & intMID
set objRS = objConn.Execute(sqlQuery)
' �ڷ��� �ʵ� ������ �д´�. 
strSubject = objRS("MSubject")
' �Խ��� ������ CR ���ڸ� "<BR>" �� �ٲ۴�. 
strContent = Replace(objRS("MContent"), chr(13), "<BR>")
strUName = objRS("MUName")
strUFName = objRS("MUFName")
strUFSize = objRS("MUFSize")
%> 
<HTML>
<HEAD> 
<TITLE> JaRyo View </TITLE>
</HEAD>
<BODY>
<I>�ڷẸ��</I><BR>
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
<% If strUFName <> "" Then 
    strUURL = strMyURL & strVirUDir & "/"
    strUURL = strUURL & strUFName
%>
<TR>
    <TD COLSPAN=2 ALIGN="right"> ÷��ȭ�� : 
        <A HREF="<%= strUURL %>"><%= strUFName %></A>
        (<%= strUFSize %> Bytes) </TD>
</TR>
<% End If %>
</TABLE>
<A HREF="SJList.asp">��Ϻ���</A> 
<A HREF="SJUForm.asp?MID=<%= strMID %>">�����ϱ�</A> 
<A HREF="SJDelete.asp?MID=<%= strMID %>">�����ϱ�</A> 
</BODY>
</HTML>
<%
objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
%>
