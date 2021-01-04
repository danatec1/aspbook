<!-- #include file="SBConn.inc" -->
<%
' 프로그램의 인수를 읽는다.
strMID = Request.QueryString("MID")
' 데이터베이스에 해당 게시물을 삭제한다.
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
sqlQuery = "DELETE Message WHERE MID=" & strMID
objConn.Execute(sqlQuery) 
' 사용한 자원을 반납한다.
objConn.Close
Set objConn=Nothing
%> 
<HTML>
<HEAD>
<TITLE> Delete Message </TITLE>
</HEAD>
<BODY>
게시물이 삭제되었어요. <BR>
<A HREF="SBList.asp">목록보기</A>
</BODY>
</HTML>
