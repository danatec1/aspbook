<!-- #include file="SJConf.inc" -->
<%
' 프로그램의 인수를 읽는다.
strMID = Request.QueryString("MID")
intMID = CInt(strMID)
' 데이터베이스에서 해당 자료(레코드)를 읽는다. 
set objConn = Server.CreateObject("ADODB.Connection")
objConn.open ConnString
sqlQuery = "SELECT * FROM JMessage WHERE MID=" & intMID
set objRS = objConn.Execute(sqlQuery)
' 자료의 필드 값들을 읽는다. 
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
<I>수정하기</I>
<TABLE>
<FORM METHOD="post" ACTION="SJUpdate.asp?MID=<%= strMID %>">
<TR>
    <TD> 이름 </TD>
    <TD> <INPUT TYPE="text" NAME="frmUName" SIZE=12 MAXLENGTH=10
	  VALUE="<%= strUName %>"> </TD>
</TR> <TR>
    <TD> 제목 </TD>
    <TD> <INPUT TYPE="text" NAME="frmSubject" SIZE=55
	  VALUE="<%= strSubject %>"> </TD>
</TR> <TR>
    <TD> 내용 </TD>
    <TD> <TEXTAREA COLS="55" ROWS="5" 
	  NAME="frmContent"><%= strContent %></TEXTAREA> </TD>
</TR>
<% If strUFName <> "" Then %>
<TR>
    <TD> 파일 </TD>
    <TD> <%= strUFName %> (<%= strUFSize %> Bytes) 
</TR>
<% End If %>
</TR> <TR>       
    <TD COLSPAN=2 ALIGN="center">
	<INPUT TYPE="submit" VALUE="수정완료">
	<INPUT TYPE="reset"  VALUE="재작성">
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
