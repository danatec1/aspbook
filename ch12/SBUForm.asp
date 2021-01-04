<!-- #include file="SBConn.inc" -->
<%
' 프로그램의 인수를 읽는다.
strMID = Request.QueryString("MID")
intMID = CInt(strMID)
' 데이터베이스에서 해당 게시물(레코드)를 읽는다. 
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
sqlQuery = "SELECT * FROM Message WHERE MID=" & intMID
Set objRS = objConn.Execute(sqlQuery)
' 게시물의 필드 값들을 읽는다. 
strSubject = objRS("MSubject")
strContent = objRS("MContent")
strUName = objRS("MUName")
%> 
<HTML>
<HEAD>
<TITLE> Form for Update Message </TITLE>
</HEAD>
<BODY>
<I>수정하기</I>
<TABLE>
<FORM METHOD="post" ACTION="SBUpdate.asp?MID=<%= strMID %>">
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
</TR> <TR>       
    <TD COLSPAN=2 ALIGN="center">
	<INPUT TYPE="submit" VALUE="수정완료">
	<INPUT TYPE="reset"  VALUE="재작성">
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
