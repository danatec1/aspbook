<!-- #include file="SBConn.inc" -->
<%
' �����ͺ��̽����� �Խù����� �о�´�. 
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.open ConnString
sqlQuery = "SELECT * FROM Message ORDER BY MID DESC"
Set objRS = objConn.Execute(sqlQuery)
%> 
<HTML>
<HEAD> 
<TITLE> Message List </TITLE>
</HEAD>
<BODY>
<A HREF="SBNForm.asp">���۾���</A>
<TABLE BORDER="1">
<TR>
    <TH> ��ȣ </TH>
    <TH> ���� </TH>
    <TH> �۾��� </TH>
</TR>
<% 
' ���ڵ�¿� ���ڵ尡 ������,
IF objRS.BOF and objRS.EOF then %>
    <TR>
        <TD COLSPAN=3> �Խù��� �ϳ��� �����. <TD>
    </TR>
<% 
Else
' ���ڵ�¿� ���ڵ���� ������, �� ���ڵ徿 �ݺ��Ͽ� ����Ѵ�.
    Do Until objRS.EOF %>
	<TR>
	    <TD> <% = objRS("MID") %> </TD>
            <TD> <A HREF="SBView.asp?MID=<% = objRS("MID")%>">
	         <% = objRS("MSubject") %></A></TD>
            <TD> <% = objRS("MUName") %> </TD>
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
