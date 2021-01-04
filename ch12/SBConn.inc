<%
' 데이터베이스에 연결할 연결 문자열을 지정한다.
ConnString = "DSN=aspbook;UID=sa;PWD=;"
' 회원인지를 체크한다.
strLogin = Session("sesLogin")
If (strLogin = "") Then
    Response.Redirect "Login.asp"
End If
%>
