<!-- #include file="SJConf.inc" -->
<%
' �����ͺ��̽����� �ڷ���� �о�´�. 
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.open ConnString
sqlQuery = "SELECT * FROM JMessage ORDER BY MID DESC"
Set objRS = objConn.Execute(sqlQuery)
%> 
<HTML>
<HEAD> 
<TITLE> JaRyo List </TITLE>
</HEAD>
<BODY>
<A HREF="SJNForm.asp">�ڷ�ø���</A>
<TABLE BORDER="1">
<TR>
    <TH> ��ȣ </TH>
    <TH> ���� </TH>
    <TH> ���� </TH>
    <TH> �۾��� </TH>
</TR>
<% 
' ���ڵ�¿� ���ڵ尡 ������,
IF objRS.BOF and objRS.EOF then %>
    <TR>
        <TD COLSPAN=4> �ڷᰡ �ϳ��� �����. <TD>
    </TR>
<% 
Else
' ���ڵ�¿� ���ڵ���� ������, �� ���ڵ徿 �ݺ��Ͽ� ����Ѵ�.
    Do Until objRS.EOF 
	strMID = objRS("MID")
        strSubject = objRS("MSubject")
        strUName = objRS("MUName")
        strUFName = objRS("MUFName")
%>
	<TR>
	    <TD> <% = strMID %> </TD>
<%      If ( strUFName <> "") Then %>
	    <TD> @ </TD>
<%      Else %>
	    <TD> &nbsp; </TD>
<%      End If %>
            <TD> <A HREF="SJView.asp?MID=<% = strMID %>">
	         <% = strSubject %></A></TD>
            <TD> <% = strUName %> </TD>
        </TR>
<%      objRS.MoveNext
    Loop 
End If %>
</TABLE>
</BODY>
</HTML>
<%
' ����� �ڿ��� �ݳ��Ѵ�.
objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
%>
