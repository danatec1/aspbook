<!-- #include file="SBConn.inc" -->
<%
' ���α׷��� �μ��� �д´�.
strMID = Request.QueryString("MID")
' �����ͺ��̽��� �ش� �Խù��� �����Ѵ�.
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
sqlQuery = "DELETE Message WHERE MID=" & strMID
objConn.Execute(sqlQuery) 
' ����� �ڿ��� �ݳ��Ѵ�.
objConn.Close
Set objConn=Nothing
%> 
<HTML>
<HEAD>
<TITLE> Delete Message </TITLE>
</HEAD>
<BODY>
�Խù��� �����Ǿ����. <BR>
<A HREF="SBList.asp">��Ϻ���</A>
</BODY>
</HTML>
