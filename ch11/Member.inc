<%
' 회원으로 Login 을 하였는지를 체크한다.
strLogin = Session("sesLogin")
If (strLogin = "") Then
' Login을 하지 않았으면, Login 절차로 되돌린다.
    Response.Redirect "Login.asp"
End If
' Login을 하였다. 그러므로, 현재의 웹 페이지를 서비스한다.
%>
