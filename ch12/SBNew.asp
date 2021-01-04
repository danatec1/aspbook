<!-- #include file="SBConn.inc" -->
<%
' 사용자의 입력들을 읽는다.
strUName = Request.Form("frmUName")
strSubject = Request.Form("frmSubject")
strContent = Request.Form("frmContent")
' 단일(') 인용 부호 문제를 없앤다. 
strSubject = Replace(strSubject, "'", "''")
strContent = Replace(strContent, "'", "''")
' 데이터베이스에 데이터들을 저장한다.
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
sqlQuery = "INSERT INTO Message VALUES (" 
sqlQuery = sqlQuery & "'" & strSubject & "'" 
sqlQuery = sqlQuery & ",'" & strContent & "'"
sqlQuery = sqlQuery & ",'" & strUName & "')"
objConn.Execute(sqlQuery) 
%> 
<HTML>
<HEAD>
<TITLE> New Message </TITLE>
</HEAD>
<BODY>
게시물이 등록되었어요. <BR>
<A HREF="SBList.asp">목록보기</A>
</BODY>
</HTML>
<%
objConn.Close
Set objConn = Nothing
%>
