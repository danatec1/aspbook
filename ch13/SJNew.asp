<!-- #include file="SJConf.inc" -->
<%
' ������� �Էµ��� �д´�.
Set objFUL = Server.CreateObject("FileMan.FileUpLoad")
strUName = objFUL.GetValueByName("frmUName")
strSubject = objFUL.GetValueByName("frmSubject")
strContent = objFUL.GetValueByName("frmContent")
strLPFile = objFUL.GetValueByName("frmLPFile")
' ����(') �ο� ��ȣ ������ ���ش�. 
strSubject = Replace(strSubject, "'", "''")
strContent = Replace(strContent, "'", "''")
'
' Upload ������ ������  
If strLPFile = "" Then
'   ���� ������ ������ �����͵��� �����ͺ��̽��� �����Ѵ�.
    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.Open ConnString
    sqlQuery = "INSERT INTO JMessage VALUES (" 
    sqlQuery = sqlQuery & "'" & strSubject & "'" 
    sqlQuery = sqlQuery & ",'" & strContent & "'"
    sqlQuery = sqlQuery & ",'" & strUName & "'"
    sqlQuery = sqlQuery & ",'','')" 
    objConn.Execute(sqlQuery) 
    objConn.Close
    Set objConn = Nothing
    Set objFUL = Nothing
%>
    <HTML>
    <HEAD>
    <TITLE> New JaRyo </TITLE>
    </HEAD>
    <BODY>
    �ڷᰡ ��ϵǾ����. <BR>
    ÷�� ������ �����ϴ�. <BR>
    <A HREF="SJList.asp">��Ϻ���</A>
    </BODY>
    </HTML>
<%
'
' Upload ������ ������  
Else
'   ������ ũ�⸦ �д´�.
    strUFSize=objFUL.TotalBytes("frmLPFile")
    strUFSize=CStr(strUFSize)
'   ���� �̸����� �����Ѵ�.
    strUFName = Mid(strLPFile, InStrRev(strLPFile,"\")+1)
    strUPFile = strPhyUDir & "/" & strUFName
'   On Error Resume Next
'   ������ �����Ѵ�.
    strRet = objFUL.SaveFile("frmLPFile", CStr(strUPFile))
'   ���� UpLoad ����
    If strRet="" Then
        Set objFUL = Nothing        
%>
        <HTML>
        <HEAD>
        <TITLE> New JaRyo </TITLE>
        </HEAD>
        <BODY>
        ������ ������ ���� Error�� �߻��߾��.<BR>
        <%=strUFName%>�� ������ �̹� �ְų� <BR>
        �ڷ�� ������ �̻��� �ֽ��ϴ�. <BR> 
        <A HREF="SJList.asp">��Ϻ���</A>
        </BODY>
        </HTML>
<%
'   ���� UpLoad ����
    Else 
'       �����ͺ��̽��� �����͵��� �����Ѵ�.
        Set objConn = Server.CreateObject("ADODB.Connection")
        objConn.Open ConnString
        sqlQuery = "INSERT INTO JMessage VALUES (" 
        sqlQuery = sqlQuery & "'" & strSubject & "'" 
        sqlQuery = sqlQuery & ",'" & strContent & "'"
        sqlQuery = sqlQuery & ",'" & strUName & "'"
        sqlQuery = sqlQuery & ",'" & strUFName & "'"
        sqlQuery = sqlQuery & ",'" & strUFSize & "')"
        objConn.Execute(sqlQuery) 
        objConn.Close
        Set objConn = Nothing
        Set objFUL = Nothing
%> 
        <HTML>
        <HEAD>
        <TITLE> New JaRyo </TITLE>
        </HEAD>
        <BODY>
        �ڷᰡ ��ϵǾ����. <BR>
        ÷�ε� ���� �̸��� <%=strUFName%>�̰�, <BR>
        ���� ũ��� <%= strUFSize%> �̴�. <BR>
        <A HREF="SJList.asp">��Ϻ���</A>
        </BODY>
        </HTML>
<%
    End If
End If
%>
