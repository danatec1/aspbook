<!-- 
#include file="YConf.inc" 
-->
<%
' ������� �Է��� �а� �μ� ��ϵ� �����.
strMID = Request.QueryString("MID")
' �ش� �Խù��� �д´�.
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
sqlQuery = "SELECT * FROM " & strTable 
sqlQuery = sqlQuery & " WHERE MID =" & strMID
Set objRS = objConn.Execute(SQLQuery)
' ��ȣ�� �д´�.
strUPassword = Request.Form("MUPassword")
If ( objRS("MUPassword") = strUPassword) then
'   �ش� �Խù��� �����Ѵ�.
    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.Open ConnString
    sqlQuery = "DELETE " & strTable 
    sqlQuery = sqlQuery & " WHERE MID =" & strMID
    objConn.Execute(sqlQuery)
    objConn.Close
    Set objConn = Nothing
    strURL = "YMList.asp?" & CArg3
    Response.Redirect strURL
Else
%>
    <HTML>
    <HEAD>
    <TITLE> Password Error </TITLE>
    </HEAD>
    <BODY>
    ��ȣ�� Ʋ���ϴ�. <BR>
    <A HREF="javascript:history.back()">�ǵ��ư���</A>
    </BODY>
    </HTML>
<%
End If
%>
