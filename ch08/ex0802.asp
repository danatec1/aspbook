<%
' ������� Login�� Password�� �д´�.
strLogin = Request.Form("frmLogin")
strPassword = Request.Form("frmPassword")
' Login�� Check �Ѵ�.
If (strLogin <> "guest") Then
    Session("sesStatus") = 1
    Response.Redirect "ex0801.asp"
End If
' Password�� Check �Ѵ�.
If (strPassword <> "guest") Then
    Session("sesStatus") = 2
    Response.Redirect "ex0801.asp"
End If
' Login�� Password�� ��Ȯ�ϸ� ������ ���� ����Ѵ�.
Session("sesLogin") = strLogin
Session("sesPassword") = strPassword
' ���� ���ø����̼��� �����Ѵ�.
Response.Redirect "ex0803.asp"
%>
