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
strContent = objRS("MContent")
strUName = objRS("MUName")
%> 
<HTML>
<HEAD>
<TITLE> Form for Update Message </TITLE>
</HEAD>
<BODY>
<I>�����ϱ�</I>
<TABLE>
<FORM METHOD="post" ACTION="SBUpdate.asp?MID=<%= strMID %>">
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
</TR> <TR>       
    <TD COLSPAN=2 ALIGN="center">
	<INPUT TYPE="submit" VALUE="�����Ϸ�">
	<INPUT TYPE="reset"  VALUE="���ۼ�">
    </TD>
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
