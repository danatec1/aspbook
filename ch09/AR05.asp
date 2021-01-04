<%
' 로그 파일에 기록할 시간과 URL을 구한다.
strNow = Now()
adHomePageURL = Request.QueryString("url")
strLog = strNow & "  " & adHomePageURL
' 로그 파일에 시간과 URL을 기록한다.
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
strPath = Server.MapPath("/booksrc/ch09")
strPFile = strPath & "/" & "ARLog.txt"
' 파일에 내용을 첨부(8)하기 위해서 열며,
' 파일이 존재하지 않으면 새로이 생성(True)한다.
Set objTS = objFSO.OpenTextFile(strPFile, 8, True)
objTS.WriteLine(strLog)
objTS.Close
Set objTS = Nothing
' 광고주의 홈페이지로 이동한다.
Response.Redirect adHomePageURL
%>
