<%
' 사용자의 입력을 읽는다.
strLogin = Request.Form("frmLogin")
strPassword = Request.Form("frmPassword")
' 데이터베이스에 연결하여 사용자 입력 계정을 읽는다.
ConnString = "DSN=aspbook;UID=sa;PWD=;"
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
sqlQuery = "SELECT * FROM Account " 
sqlQuery = sqlQuery & " WHERE login='" & strLogin & "'" 
Set objRS = objConn.Execute(sqlQuery)
' 계정이 없으면,
If (objRS.BOF and objRS.EOF) Then
    Response.Redirect "Login.asp?Status=1"
End If
' 계정이 있으면, 암호가 같은 지를 체크한다.
If (strPassword <> objRS("password")) Then
    Response.Redirect "Login.asp?Status=2"
End If
' 암호가 같으면, 응용 프로그램으로 들어간다.
Session("sesLogin") = strLogin
Response.Redirect "SBList.asp"
%>
