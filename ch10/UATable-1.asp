<%
' DSN�� ������� ���� ���� ���ڿ� 
ConnString = "Driver={SQL Server}; Server=young2; Database=aspbook;"
ConnString = ConnString & "UID=sa; PWD=;"
' Connection ��ü�� ������ �Ŀ� ������ ����ҿ� �����Ѵ�.
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
' ������ �߻��ص� ���� ������ ��� ����ǵ��� �Ѵ�.
on Error Resume Next
' Connection ��ü�� ����Ͽ� ����� �����Ѵ�.
sqlQuery = " CREATE TABLE Account ( "
sqlQuery = sqlQuery & " login      Varchar(10)  NOT NULL,"
sqlQuery = sqlQuery & " password  Varchar(10)  NOT NULL,"
sqlQuery = sqlQuery & " email      Varchar(60)  NULL)"
objConn.Execute sqlQuery
%>
<HTML>
<HEAD>
<TITLE> User Account Table Creation </TITLE>
</HEAD>
<BODY>
<%
If (objConn.Errors.Count > 0) Then
    For each Error in objConn.Errors
        Response.Write("Error #:" & Error.Number & "<BR>")
        Response.Write("Error Description:" & Error.Description)
    Next
Else
%>
ȸ�� ���� ���̺��� �����Ǿ����.
<%
End If
%>
</BODY>
</HTML>
<%
objConn.Close
Set objConn = Nothing
%>
