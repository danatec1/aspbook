<%
' DSN을 사용하지 않은 연결 문자열 
ConnString = "Driver={SQL Server}; Server=young2; Database=aspbook;"
ConnString = ConnString & "UID=sa; PWD=;"
' Connection 객체를 생성한 후에 데이터 저장소에 연결한다.
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
' 오류가 발생해도 다음 문에서 계속 실행되도록 한다.
on Error Resume Next
' Connection 객체를 사용하여 명령을 실행한다.
sqlQuery = " CREATE TABLE Account ( "
sqlQuery = sqlQuery & " login      Varchar(10)  NOT NULL,"
sqlQuery = sqlQuery & " password  Varchar(10)  NOT NULL,"
sqlQuery = sqlQuery & " email      Varchar(60)  NULL)"
objConn.Execute sqlQuery
%>
<HTML>
<HEAD>
<TITLE> User Account Table Creation </TITLE>
</HEAD>
<BODY>
<%
If (objConn.Errors.Count > 0) Then
    For each Error in objConn.Errors
        Response.Write("Error #:" & Error.Number & "<BR>")
        Response.Write("Error Description:" & Error.Description)
    Next
Else
%>
회원 계정 테이블이 생성되었어요.
<%
End If
%>
</BODY>
</HTML>
<%
objConn.Close
Set objConn = Nothing
%>
