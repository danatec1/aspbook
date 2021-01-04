<%
On Error Resume Next 
' 데이터 파일의 물리적 디렉토리 이름을 찾는다.
strPath = Server.MapPath("/booksrc/ch09")
strPFile = strPath & "/" & "gbook.txt"
' FileSystemObject 객체를 생성한다.
Set objFSO = CreateObject("Scripting.FileSystemObject")
' 파일에 내용을 읽기(1) 위해서 Open 한다.
' 만일, 파일이 존재하지 않으면 새로이 생성(True)한다.
Set objTS = objFSO.OpenTextFile(strPFile, 1, True)
%>
<HTML>
<HEAD>
<TITLE> Guest Book : List </TITLE>
</HEAD>
<BODY BGCOLOR=white>
<CENTER>
<H2> 방명록 </H2>
<A HREF="GBNForm.asp">글쓰기</A>
<TABLE>
<% 
If objTS.AtEndOfStream Then %>
<TABLE BORDER=1 WIDTH=400>
    <TR>
        <TD> 방명록에 읽을 글이 없습니다. </TD>
    </TR>
</TABLE>
<%
Else 
    Do Until objTS.AtEndOfStream %>
        <TABLE BORDER=1 WIDTH=400>
        <TR>
            <TD WIDTH=50> 이름 </TD> <TD> <% = objTS.ReadLine %> </TD>
        </TR> <TR> 
            <TD WIDTH=50> 내용 </TD> <TD> <% = objTS.ReadLine %> </TD>
        </TR> <TR> 
            <TD WIDTH=50> 일시 </TD> <TD> <% = objTS.ReadLine %> </TD>
        </TR>
        </TABLE>
	<BR>
<%  Loop
End If
' 객체 사용을 중단하고 파기한다.
objTS.Close
Set objTS = Nothing
%>
</CENTER>
</BODY>
</HTML>
