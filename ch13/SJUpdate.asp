<!-- #include file="SJConf.inc" -->
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
sqlQuery = "UPDATE JMessage SET " 
sqlQuery = sqlQuery & " MSubject='" & strSubject & "'" 
sqlQuery = sqlQuery & ",MContent='" & strContent & "'"
sqlQuery = sqlQuery & ",MUName='" & strUName & "'"
sqlQuery = sqlQuery & " WHERE MID=" & strMID
objConn.Execute(sqlQuery) 
%> 
<HTML>
<HEAD>
<TITLE> Update JaRyo </TITLE>
</HEAD>
<BODY>
�ڷ� ������ �����Ǿ����. <BR>
<A HREF="SJList.asp">��Ϻ���</A>
</BODY>
</HTML>
<%
objConn.Close
Set objConn = Nothing
%>
