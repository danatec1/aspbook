<%
' 입력 데이터들을 읽는다.
strLogin = Request.Form("frmLogin")
strPassword = Request.Form("frmPassword")
strEmail = Request.Form("frmEmail")
' DSN을 사용한 연결 문자열  
ConnString = "DSN=aspbook;UID=sa;PWD=;"
' Connection 객체를 생성한 후에 데이터 저장소에 연결한다.
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
' Connection 객체를 사용하여 명령을 실행한다.
sqlQuery = " INSERT INTO Account VALUES ("
sqlQuery = sqlQuery  &  "'"  &  strLogin   &  "'"
sqlQuery = sqlQuery  &  ",'"  &  strPassword  &  "'"
sqlQuery = sqlQuery  &  ",'"  &  strEmail   &  "')"
objConn.Execute sqlQuery
%>
<HTML>
<HEAD>
<TITLE> User Account Registration </TITLE>
</HEAD>
<BODY>
회원 계정이 등록되었어요.
</BODY>
</HTML>
<%
objConn.Close
Set objConn = Nothing
%>
