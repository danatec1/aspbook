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
' 게시판 내용의 CR 문자를 "<BR>" 로 바꾼다. 
strContent = Replace(objRS("MContent"), chr(13), "<BR>")
strUName = objRS("MUName")
%> 
<HTML>
<HEAD> 
<TITLE> Message View </TITLE>
</HEAD>
<BODY>
<I>글읽기</I><BR>
<TABLE border=1>
<TR>
    <TD> 이름 </TD>
    <TD> <%= strUName %> </TD>	
</TR> <TR>
    <TD> 제목 </TD>
    <TD> <%= strSubject %> </TD>
</TR> <TR>
    <TD> 내용 </TD>
    <TD> <%= strContent %> </TD>	
</TR>
</TABLE>
<A HREF="SBList.asp">목록보기</A> 
<A HREF="SBUForm.asp?MID=<%= strMID %>">수정하기</A> 
<A HREF="SBDelete.asp?MID=<%= strMID %>">삭제하기</A> 
</BODY>
</HTML>
<%
objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
%>
