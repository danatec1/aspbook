<%
On Error Resume Next 
' FileSystemObject ��ü�� �����Ѵ�.
Set objFSO = CreateObject("Scripting.FileSystemObject")
' ���Ͽ� ������ ÷��(8)�ϱ� ���ؼ� Open �Ѵ�.
Set objTS = objFSO.OpenTextFile("C:\test.txt", 8)
%>
<HTML>
<HEAD>
<TITLE> Example : TextStream Object </TITLE>
</HEAD>
<BODY BGCOLOR=white>
<CENTER>
<% 
datNow = now()
' date ���� ���������� �ٲ۴�.
strNow = CStr(datNow) 
msg = "������ " & strNow & "�Դϴ�."
If IsObject(objTS) Then 
    objTS.WriteLine(msg)
    objTS.Close()
%>
    (<%= msg %>) �޼����� <BR>
    C:\test.txt ���Ͽ� ����Ͽ����ϴ�.
<% 
Else %>
    C:\test.txt ������ �����. 
<% 
End If %>
</CENTER>
</BODY>
</HTML>
