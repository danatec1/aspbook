<%
' ȸ������ Login �� �Ͽ������� üũ�Ѵ�.
strLogin = Session("sesLogin")
If (strLogin = "") Then
' Login�� ���� �ʾ�����, Login ������ �ǵ�����.
    Response.Redirect "Login.asp"
End If
' Login�� �Ͽ���. �׷��Ƿ�, ������ �� �������� �����Ѵ�.
%>