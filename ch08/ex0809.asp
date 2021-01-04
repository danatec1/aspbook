<% 
On Error Resume Next
' FileSystemObject 객체를 생성한다.
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
' 파일을 오픈하고, TextStream 객체를 리턴한다.
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
    C:\test.txt 파일이 Open되었어요.
<% 
Else %>
    C:\test.txt 파일이 없어요. 
<% 
End If %> 
</CENTER>
</BODY>
</HTML>
