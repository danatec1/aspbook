<% 
' FileSystemObject ��ü�� �����Ѵ�.
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
' ���ο� ���� test.txt�� �����Ѵ�. 
Set objTS = objFSO.CreateTextFile("C:\test.txt", True)
%>
<HTML>
<HEAD>
<TITLE> Example : CreateTextFile Method</TITLE>
</HEAD>
<BODY BGCOLOR=white>
<CENTER>
C:\test.txt ������ ��������. <BR>
Ž������� �̿��Ͽ� ������ �����Ǿ������� Ȯ���ϼ���.
</CENTER>
</BODY>
</HTML>
