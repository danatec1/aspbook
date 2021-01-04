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
' 게시판 내용의 CR 문자를 "<BR>" 로 바꾼다. 
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
<I>자료보기</I><BR>
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
<% If strUFName <> "" Then 
    strUURL = strMyURL & strVirUDir & "/"
    strUURL = strUURL & strUFName
%>
<TR>
    <TD COLSPAN=2 ALIGN="right"> 첨부화일 : 
        <A HREF="<%= strUURL %>"><%= strUFName %></A>
        (<%= strUFSize %> Bytes) </TD>
</TR>
<% End If %>
</TABLE>
<A HREF="SJList.asp">목록보기</A> 
<A HREF="SJUForm.asp?MID=<%= strMID %>">수정하기</A> 
<A HREF="SJDelete.asp?MID=<%= strMID %>">삭제하기</A> 
</BODY>
</HTML>
<%
objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
%>
