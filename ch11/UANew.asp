<%
' �Է� �����͵��� �д´�.
strLogin = Request.Form("frmLogin")
strPassword = Request.Form("frmPassword")
strEmail = Request.Form("frmEmail")
' DSN�� ����� ���� ���ڿ�  
ConnString = "DSN=aspbook;UID=sa;PWD=;"
' Connection ��ü�� ������ �Ŀ� ������ ����ҿ� �����Ѵ�.
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
' Connection ��ü�� ����Ͽ� ����� �����Ѵ�.
sqlQuery = " INSERT INTO Account VALUES ("
sqlQuery = sqlQuery  &  "'"  &  strLogin   &  "'"
sqlQuery = sqlQuery  &  ",'"  &  strPassword  &  "'"
sqlQuery = sqlQuery  &  ",'"  &  strEmail   &  "')"
objConn.Execute sqlQuery
%>
<HTML>
<HEAD>
<TITLE> User Account Registration </TITLE>
</HEAD>
<BODY>
ȸ�� ������ ��ϵǾ����.
</BODY>
</HTML>
<%
objConn.Close
Set objConn = Nothing
%>
