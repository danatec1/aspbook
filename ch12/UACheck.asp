<%
' ������� �Է��� �д´�.
strLogin = Request.Form("frmLogin")
strPassword = Request.Form("frmPassword")
' �����ͺ��̽��� �����Ͽ� ����� �Է� ������ �д´�.
ConnString = "DSN=aspbook;UID=sa;PWD=;"
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
sqlQuery = "SELECT * FROM Account " 
sqlQuery = sqlQuery & " WHERE login='" & strLogin & "'" 
Set objRS = objConn.Execute(sqlQuery)
' ������ ������,
If (objRS.BOF and objRS.EOF) Then
    Response.Redirect "Login.asp?Status=1"
End If
' ������ ������, ��ȣ�� ���� ���� üũ�Ѵ�.
If (strPassword <> objRS("password")) Then
    Response.Redirect "Login.asp?Status=2"
End If
' ��ȣ�� ������, ���� ���α׷����� ����.
Session("sesLogin") = strLogin
Response.Redirect "SBList.asp"
%>
