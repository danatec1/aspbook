<!-- #include file="SJConf.inc" -->
<%
' 프로그램의 인수를 읽는다.
strMID = Request.QueryString("MID")
' 데이터베이스에 연결하여 해당 레코드를 읽는다.
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
sqlQuery = "SELECT * FROM JMessage WHERE MID=" & strMID 
Set objRS = objConn.Execute(sqlQuery) 
' 첨부 파일 이름을 읽는다.
strUFName = objRS("MUFName")
' 첨부 파일이 있으면, 파일을 먼저 삭제한다.  
If (strUFName <> "") Then
    strUPFile = strPhyUDir & "/" & strUFName
    On Error Resume Next
    Set objFSO=CreateObject("Scripting.FileSystemObject")
    objFSO.DeleteFile(strUPFile)
    Set objFSO=Nothing
End If
' 데이터베이스 테이블에 해당 자료(레코드)를 삭제한다.
sqlQuery = "DELETE JMessage WHERE MID=" & strMID
objConn.Execute(sqlQuery) 
' 사용한 자원을 반납한다.
objRS.Close 
Set objRS = Nothing
objConn.Close
Set objConn=Nothing
%> 
<HTML>
<HEAD>
<TITLE> Delete JaRyo </TITLE>
</HEAD>
<BODY>
자료가 삭제되었어요. <BR>
<A HREF="SJList.asp">목록보기</A>
</BODY>
</HTML>
