<!-- #include file="SBConn.inc" -->
<%
' ������� �Էµ��� �д´�.
strUName = Request.Form("frmUName")
strSubject = Request.Form("frmSubject")
strContent = Request.Form("frmContent")
' ����(') �ο� ��ȣ ������ ���ش�. 
strSubject = Replace(strSubject, "'", "''")
strContent = Replace(strContent, "'", "''")
' �����ͺ��̽��� �����͵��� �����Ѵ�.
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
sqlQuery = "INSERT INTO Message VALUES (" 
sqlQuery = sqlQuery & "'" & strSubject & "'" 
sqlQuery = sqlQuery & ",'" & strContent & "'"
sqlQuery = sqlQuery & ",'" & strUName & "')"
objConn.Execute(sqlQuery) 
%> 
<HTML>
<HEAD>
<TITLE> New Message </TITLE>
</HEAD>
<BODY>
�Խù��� ��ϵǾ����. <BR>
<A HREF="SBList.asp">��Ϻ���</A>
</BODY>
</HTML>
<%
objConn.Close
Set objConn = Nothing
%>
