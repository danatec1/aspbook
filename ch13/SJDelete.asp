<!-- #include file="SJConf.inc" -->
<%
' ���α׷��� �μ��� �д´�.
strMID = Request.QueryString("MID")
' �����ͺ��̽��� �����Ͽ� �ش� ���ڵ带 �д´�.
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
sqlQuery = "SELECT * FROM JMessage WHERE MID=" & strMID 
Set objRS = objConn.Execute(sqlQuery) 
' ÷�� ���� �̸��� �д´�.
strUFName = objRS("MUFName")
' ÷�� ������ ������, ������ ���� �����Ѵ�.  
If (strUFName <> "") Then
    strUPFile = strPhyUDir & "/" & strUFName
    On Error Resume Next
    Set objFSO=CreateObject("Scripting.FileSystemObject")
    objFSO.DeleteFile(strUPFile)
    Set objFSO=Nothing
End If
' �����ͺ��̽� ���̺� �ش� �ڷ�(���ڵ�)�� �����Ѵ�.
sqlQuery = "DELETE JMessage WHERE MID=" & strMID
objConn.Execute(sqlQuery) 
' ����� �ڿ��� �ݳ��Ѵ�.
objRS.Close 
Set objRS = Nothing
objConn.Close
Set objConn=Nothing
%> 
<HTML>
<HEAD>
<TITLE> Delete JaRyo </TITLE>
</HEAD>
<BODY>
�ڷᰡ �����Ǿ����. <BR>
<A HREF="SJList.asp">��Ϻ���</A>
</BODY>
</HTML>
