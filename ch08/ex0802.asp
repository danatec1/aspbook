<%
' 사용자의 Login과 Password를 읽는다.
strLogin = Request.Form("frmLogin")
strPassword = Request.Form("frmPassword")
' Login을 Check 한다.
If (strLogin <> "guest") Then
    Session("sesStatus") = 1
    Response.Redirect "ex0801.asp"
End If
' Password를 Check 한다.
If (strPassword <> "guest") Then
    Session("sesStatus") = 2
    Response.Redirect "ex0801.asp"
End If
' Login과 Password가 정확하면 데이터 값을 기억한다.
Session("sesLogin") = strLogin
Session("sesPassword") = strPassword
' 실제 어플리케이션을 실행한다.
Response.Redirect "ex0803.asp"
%>
