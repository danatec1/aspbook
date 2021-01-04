<%
' DSN을 사용하지 않은 연결 문자열 
ConnString = "Driver={SQL Server}; Server=young2; Database=aspbook;"
ConnString = ConnString & "UID=sa; PWD=;"
' Connection 객체를 생성한 후에 데이터 저장소에 연결한다. 
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
' Connection 객체를 사용하여 명령을 실행한다.
sqlQuery = " CREATE TABLE Account ( "
sqlQuery = sqlQuery & " login      Varchar(10)  NOT NULL,"
sqlQuery = sqlQuery & " password  Varchar(10)  NOT NULL,"
sqlQuery = sqlQuery & " email      Varchar(60)  NULL)"
objConn.Execute sqlQuery
%>
<HTML>
<HEAD>
<TITLE> User Account Table Creation</TITLE>
</HEAD>
<BODY>
회원 계정 테이블이 생성되었어요.
</BODY>
</HTML>
<%
objConn.Close
Set objConn = Nothing
%>

