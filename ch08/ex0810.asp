<%
On Error Resume Next 
' FileSystemObject 객체를 생성한다.
Set objFSO = CreateObject("Scripting.FileSystemObject")
' 파일에 내용을 첨부(8)하기 위해서 Open 한다.
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
' date 형을 문자형으로 바꾼다.
strNow = CStr(datNow) 
msg = "지금은 " & strNow & "입니다."
If IsObject(objTS) Then 
    objTS.WriteLine(msg)
    objTS.Close()
%>
    (<%= msg %>) 메세지를 <BR>
    C:\test.txt 파일에 기록하였습니다.
<% 
Else %>
    C:\test.txt 파일이 없어요. 
<% 
End If %>
</CENTER>
</BODY>
</HTML>
