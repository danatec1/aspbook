<%
Dim intNumRecs
' ������� �Է��� �д´�.
strLogin = Request.Form("frmLogin")
' Command ��ü�� �����Ѵ�.
Set objComm = Server.CreateObject("ADODB.Command")
' ActiveConnection �Ӽ��� ���� ���ڿ��� �����Ѵ�.
ConnString = "DSN=aspbook;UID=sa;PWD=;"
objComm.ActiveConnection = ConnString
' CommandText �Ӽ��� ������ ���Ǿ �����Ѵ�.
sqlQuery = " DELETE FROM Account WHERE login='" & strLogin & "'"
objComm.CommandText = sqlQuery 
' ���ǹ��� �����Ͽ� ������ �����Ѵ�.
objComm.Execute intNumRecs
%>
<HTML>
<HEAD>
<TITLE> User Account Delete </TITLE>
</HEAD>
<BODY>
<% 
If (intNumRecs > 0) Then %>
    ȸ�� ������ �����߾��.
<% 
Else %>
    ȸ�� ������ �����.
<% 
End If %>
</BODY>
</HTML>
<%
Set objComm = Nothing
%>
