<%
' �����ͺ��̽��� ���� ���ڿ��� �����Ѵ�.
ConnString = "DSN=aspbook;UID=sa;PWD=;"
' Connection ��ü�� �����ϰ� �����ͺ��̽��� �����Ѵ�.
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
' ù��° ����
sqlQuery = "CREATE TABLE TTest (GCODE Varchar(10)  NOT NULL)"
objConn.Execute sqlQuery
' �ι�° ����
sqlQuery = "INSERT INTO TTT VALUES ('A1001') "
objConn.Execute sqlQuery
%>
<HTML>
<HEAD>
<TITLE> Table Creation and Data Insert </TITLE>
</HEAD>
<BODY>
���̺��� �����ϰ� �����͸� �Է��߾��.
</BODY>
</HTML>
<%
objConn.Close
Set objConn = Nothing
%>
