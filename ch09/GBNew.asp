<%
' ������� �Է��� �д´�.
strName = Request.Form("frmName")
strContent = Request.Form("frmContent")
' ���Ͽ� �����͵��� ���� ������ �����ϱ� ������
' [Enter] Ű�� �����Ǵ� \r\n ������ ó���Ѵ�.   
strContent = Replace(strContent, chr(13), "&nbsp;")
strContent = Replace(strContent, chr(10), "<BR>")
' ������ �Ͻ� ������ ã�´�.
datNow = Now()
' date ���� ���������� �ٲ۴�.
strNow = CStr(datNow) 
' ������ ������ ������ ���丮 �̸��� ã�´�.
strPath = Server.MapPath("/booksrc/ch09")
strPFile = strPath & "/" & "gbook.txt"
' FileSystemObject ��ü�� �����Ѵ�.
Set objFSO = CreateObject("Scripting.FileSystemObject")
' ���Ͽ� ������ ÷��(8)�ϱ� ���ؼ� Open �Ѵ�.
Set objTS = objFSO.OpenTextFile(strPFile, 8)
' �����͵��� ���Ͽ� �����Ѵ�.
objTS.WriteLine strName  
objTS.WriteLine strContent
objTS.WriteLine strNow 
'
objTS.Close
Set objTS = Nothing
%> 
<HTML>
<HEAD>
<TITLE> Guest Book : New </TITLE>
</HEAD>
<BODY>
�����մϴ�. ���� ��ϵǾ����. <BR>
<A HREF="GBList.asp">����</A>
</BODY>
</HTML>
