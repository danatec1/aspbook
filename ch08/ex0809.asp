<% 
On Error Resume Next
' FileSystemObject ��ü�� �����Ѵ�.
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
' ������ �����ϰ�, TextStream ��ü�� �����Ѵ�.
Set objTS = objFSO.OpenTextFile("C:\test.txt")
%>
<HTML>
<HEAD>
<TITLE> Example : OpenTextFile Method </TITLE>
</HEAD>
<BODY BGCOLOR=white>
<CENTER>
<% 
If IsObject(objTS) Then %>
    C:\test.txt ������ Open�Ǿ����.
<% 
Else %>
    C:\test.txt ������ �����. 
<% 
End If %> 
</CENTER>
</BODY>
</HTML>
