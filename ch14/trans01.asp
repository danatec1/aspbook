<%
' 데이터베이스의 연결 문자열을 지정한다.
ConnString = "DSN=aspbook;UID=sa;PWD=;"
' Connection 객체를 생성하고 데이터베이스에 연결한다.
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
' 첫번째 연산
sqlQuery = "CREATE TABLE TTest (GCODE Varchar(10)  NOT NULL)"
objConn.Execute sqlQuery
' 두번째 연산
sqlQuery = "INSERT INTO TTT VALUES ('A1001') "
objConn.Execute sqlQuery
%>
<HTML>
<HEAD>
<TITLE> Table Creation and Data Insert </TITLE>
</HEAD>
<BODY>
테이블을 생성하고 데이터를 입력했어요.
</BODY>
</HTML>
<%
objConn.Close
Set objConn = Nothing
%>
