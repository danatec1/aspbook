<%
' 사용자의 입력을 읽는다.
strName = Request.Form("frmName")
strContent = Request.Form("frmContent")
' 파일에 데이터들을 라인 단위로 저장하기 때문에
' [Enter] 키에 대응되는 \r\n 문제를 처리한다.   
strContent = Replace(strContent, chr(13), "&nbsp;")
strContent = Replace(strContent, chr(10), "<BR>")
' 현재의 일시 정보를 찾는다.
datNow = Now()
' date 형을 문자형으로 바꾼다.
strNow = CStr(datNow) 
' 데이터 파일의 물리적 디렉토리 이름을 찾는다.
strPath = Server.MapPath("/booksrc/ch09")
strPFile = strPath & "/" & "gbook.txt"
' FileSystemObject 객체를 생성한다.
Set objFSO = CreateObject("Scripting.FileSystemObject")
' 파일에 내용을 첨부(8)하기 위해서 Open 한다.
Set objTS = objFSO.OpenTextFile(strPFile, 8)
' 데이터들을 파일에 저장한다.
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
감사합니다. 글이 등록되었어요. <BR>
<A HREF="GBList.asp">방명록</A>
</BODY>
</HTML>
