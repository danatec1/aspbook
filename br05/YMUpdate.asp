<!-- 
#include file="YConf.inc" 
-->
<%
' ������� �Է��� �а� �μ� ��ϵ� �����.
strMID = Request.QueryString("MID")
CArg4 = CArg3 & "&MID=" & strMID
' �ش� �Խù��� �д´�.
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
sqlQuery = "SELECT * FROM " & strTable 
sqlQuery = sqlQuery & " WHERE MID =" & strMID
Set objRS = objConn.Execute(SQLQuery)
'
strUPassword = Request.Form("MUPassword")
If ( objRS("MUPassword") = strUPassword) then
' ����ڵ��� �����͵��� �д´�.
    strURealName = Request.Form("MURealName")
    strSubject = Request.Form("MSubject")
    strSubject = Replace(strSubject, "'", "''")
    strContent = Request.Form("MContent")
    strContent = Replace(strContent, "'", "''")
' �ش� ���ڵ带 �����Ѵ�.
    sqlQuery = "UPDATE " & strTable 
    sqlQuery = sqlQuery & " SET MURealName='" & strURealName & "'," 
    sqlQuery = sqlQuery & "     MSubject='" & strSubject & "'," 
    sqlQuery = sqlQuery & "     MContent='" & strContent & "'" 
    sqlQUery = sqlQuery & " WHERE MID =" & strMID
    objConn.Execute(SQLQuery)
    objRS.Close
    Set objRS = Nothing
    objConn.Close
    Set objConn = Nothing
    strURL = "YMUForm.asp?" & CArg4
    Response.Redirect strURL
Else
    objRS.Close
    Set objRS=Nothing
    objConn.Close
    Set objConn=Nothing
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
