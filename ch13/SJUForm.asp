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
strContent = objRS("MContent")
strUName = objRS("MUName")
strUFName = objRS("MUFName")
strUFSize = objRS("MUFSize")
%> 
<HTML>
<HEAD>
<TITLE> Form for Update JaRyo </TITLE>
</HEAD>
<BODY>
<I>�����ϱ�</I>
<TABLE>
<FORM METHOD="post" ACTION="SJUpdate.asp?MID=<%= strMID %>">
<TR>
    <TD> �̸� </TD>
    <TD> <INPUT TYPE="text" NAME="frmUName" SIZE=12 MAXLENGTH=10
	  VALUE="<%= strUName %>"> </TD>
</TR> <TR>
    <TD> ���� </TD>
    <TD> <INPUT TYPE="text" NAME="frmSubject" SIZE=55
	  VALUE="<%= strSubject %>"> </TD>
</TR> <TR>
    <TD> ���� </TD>
    <TD> <TEXTAREA COLS="55" ROWS="5" 
	  NAME="frmContent"><%= strContent %></TEXTAREA> </TD>
</TR>
<% If strUFName <> "" Then %>
<TR>
    <TD> ���� </TD>
    <TD> <%= strUFName %> (<%= strUFSize %> Bytes) 
</TR>
<% End If %>
</TR> <TR>       
    <TD COLSPAN=2 ALIGN="center">
	<INPUT TYPE="submit" VALUE="�����Ϸ�">
	<INPUT TYPE="reset"  VALUE="���ۼ�">
    </TD>
</TR>
</FORM>
</TABLE>
</BODY>
</HTML>
<%
objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
%>
