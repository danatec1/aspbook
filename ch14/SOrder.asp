<% @ TRANSACTION=required %>
<%
' ���� �Է� ������ �д´�. 
strUName = Request.Form("UName")
strUAddress = Request.Form("UAddress")
' ���ο� ���� ��ȣ�� �����ϱ� ���ؼ� ���� ū SOID�� ��´�.
Set objConn = Session("objConn")
sqlQuery = "SELECT MAX(SOID) FROM SOrder"
Set objRS = objConn.Execute(sqlQuery)
If IsNull(objRS(0)) Then
    intSOID = 1
Else
    intSOID = 1 + objRS(0)
End If
' SUser ���̺� ����� �Է� �����Ϳ� ���� ��ȣ�� �����Ѵ�.
sqlQuery = "INSERT INTO SUser VALUES (" 
sqlQuery = sqlQuery & "'"  & strUName & "'"  
sqlQuery = sqlQuery & ",'" & strUAddress & "'"
sqlQuery = sqlQuery & ","  & intSOID & ")"
objConn.Execute(sqlQuery) 
' 
' ��ٱ����� ������ ����(Order) ���̺�� �����Ѵ�. 
'
' Shopping Cart�� ã�´�. 
strSCID = Session("sesSCID")
intSCID = CInt(strSCID)
' ��ٱ��� �����͵��� �д´�.
sqlQuery = "SELECT * FROM SCart "
sqlQuery = sqlQuery & " WHERE SCID=" & intSCID
Set objRS = objConn.Execute(sqlQuery)
' ���� ���̺�� �����͵��� �����Ѵ�.
Do Until objRS.EOF 
    strGCode = objRS("SGCODE")
    intGAmount = objRS("SGAmount")
    sqlQuery = "INSERT INTO SOrder VALUES (" 
    sqlQuery = sqlQuery        & intSOID  
    sqlQuery = sqlQuery & ",'" & strGCode & "'"
    sqlQuery = sqlQuery & ","  & intGAmount & ")"
    objConn.Execute(sqlQuery) 
    objRS.MoveNext
Loop
'
' Shopping Cart�� ����. 
'
sqlQuery = "DELETE SCart WHERE SCID=" & intSCID
objConn.Execute(sqlQuery) 
%> 
<% 
Sub OnTransactionCommit() %>
    <HTML>
    <HEAD>
    <TITLE> Transaction Success </TITLE>
    </HEAD>
    <BODY>
    ���Ű� �Ϸ�Ǿ����. ��ǰ�� �������ּż� �����մϴ�. <BR>
    <A HREF="SCView.asp">��ٱ��Ϻ���</A>
    </BODY>
    </HTML>
<%
End Sub
Sub OnTransactionAbort() %> 
    <HTML>
    <HEAD>
    <TITLE> Transaction Abort </TITLE>
    </HEAD>
    <BODY>
    ���Ű� �Ϸ���� �ʾҾ��. �ٽ� �õ����ּ���. <BR>
    <A HREF="SCView.asp">��ٱ��Ϻ���</A>
    </BODY>
    </HTML>
<%
End Sub 
%>
