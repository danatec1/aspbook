<%
' ���ε�Ǵ� ���ϵ��� ������ ���� ���丮�� �����ϰ�
'   �̿� ���� �������� ���丮�� ���Ѵ�.
strVirUPath = "/JFiles"
strPhyUPath = Server.MapPath(strVirUPath) 
'
' FileUpLoad ��ü�� ����Ͽ� �Է� �����͸� �д´�.
Set objFUL = Server.CreateObject("FileMan.FileUpLoad")
' ������ ũ�⸦ �д´�.
strUFSize=objFUL.TotalBytes("frmLPFile")
strUFSize=CStr(strUFSize)
' ���� ��θ��� ã�´�.
strLPFile = Trim(objFUL.GetValueByName("frmLPFile"))
' ���� �̸����� �����Ѵ�.
strUFName = Mid(strLPFile, InStrRev(strLPFile,"\") +1)
' ������ ������ �������� ��θ��� ���Ѵ�. 
strUPFile = strPhyUPath & "/" & strUFName
' On Error Resume Next
' ������ �����Ѵ�.
strRet = objFUL.SaveFile("frmLPFile", CStr(strUPFile))
'
' ���� UpLoad ����
If strRet="" Then
    Set objFUL = Nothing        
%>
    <HTML>
    <HEAD> </HEAD>
    <BODY>
    ������ ������ ���� Error�� �߻��߾��.<BR>
    <%=strUFName%> ������ �̹� �ְų� <BR>
    ���� ������ �̻��� �ֽ��ϴ�. <BR> 
    <A HREF="javascript:history.back()">�ǵ��ư���</A>
    </BODY>
    </HTML>
<%
' ���� UpLoad ����
Else 
%>
    <HTML>
    <HEAD></HEAD>
    <BODY>
    ������ ���������� ����Ǿ����. <BR>
    ���� �̸��� <%=strUFName%>�̰�, ũ��� <%= strUFSize%> �̴�.<BR>
    <A HREF="javascript:history.back()">�ǵ��ư���</A>
    </BODY>
    </HTML>
<%
    Set objFUL = Nothing
End If 
%>
