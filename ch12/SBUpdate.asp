<!-- #include file="SBConn.inc" -->
<%
' 프로그램의 인수를 읽는다.
strMID = Request.QueryString("MID")
' 사용자의 입력 데이터들을 읽는다.
strUName = Request.Form("frmUName")
strSubject = Request.Form("frmSubject")
strContent = Request.Form("frmContent")
' 데이터베이스에 데이터들을 수정한다.
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
sqlQuery = "UPDATE Message SET " 
sqlQuery = sqlQuery & " MSubject='" & strSubject & "'" 
sqlQuery = sqlQuery & ",MContent='" & strContent & "'"
sqlQuery = sqlQuery & ",MUName='" & strUName & "'"
sqlQuery = sqlQuery & " WHERE MID=" & strMID
objConn.Execute(sqlQuery) 
%> 
<HTML>
<HEAD>
<TITLE> Update Message </TITLE>
</HEAD>
<BODY>
게시물이 수정되었어요. <BR>
<A HREF="SBList.asp">목록보기</A>
</BODY>
</HTML>
<%
objConn.Close
Set objConn = Nothing
%>
