<%
' �����ͺ��̽��� ������ ���� ���ڿ��� �����Ѵ�.
ConnString = "DSN=aspbook;UID=sa;PWD=;"
' ȸ�������� üũ�Ѵ�.
strLogin = Session("sesLogin")
If (strLogin = "") Then
    Response.Redirect "Login.asp"
End If
%>