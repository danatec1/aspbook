<!-- #include file="SBConn.inc" -->
<%
' ���α׷��� �μ��� �д´�.
strMID = Request.QueryString("MID")
' ������� �Է� �����͵��� �д´�.
strUName = Request.Form("frmUName")
strSubject = Request.Form("frmSubject")
strContent = Request.Form("frmContent")
' �����ͺ��̽��� �����͵��� �����Ѵ�.
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
sqlQuery = "UPDATE Message SET " 
sqlQuery = sqlQuery & " MSubject='" & strSubject & "'" 
sqlQuery = sqlQuery & ",MContent='" & strContent & "'"
sqlQuery = sqlQuery & ",MUName='" & strUName & "'"
sqlQuery = sqlQuery & " WHERE MID=" & strMID
objConn.Execute(sqlQuery) 
%> 
<HTML>
<HEAD>
<TITLE> Update Message </TITLE>
</HEAD>
<BODY>
�Խù��� �����Ǿ����. <BR>
<A HREF="SBList.asp">��Ϻ���</A>
</BODY>
</HTML>
<%
objConn.Close
Set objConn = Nothing
%>
