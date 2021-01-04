<% 
' FileSystemObject 객체를 생성한다.
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
' 새로운 파일 test.txt를 생성한다. 
Set objTS = objFSO.CreateTextFile("C:\test.txt", True)
%>
<HTML>
<HEAD>
<TITLE> Example : CreateTextFile Method</TITLE>
</HEAD>
<BODY BGCOLOR=white>
<CENTER>
C:\test.txt 파일을 만들었어요. <BR>
탐색기등을 이용하여 파일이 생성되었는지를 확인하세요.
</CENTER>
</BODY>
</HTML>
