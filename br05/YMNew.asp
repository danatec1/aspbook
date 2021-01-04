<!-- 
#include file="YConf.inc" 
-->
<%
' 이름을 읽는다. 
strURealName = Trim(Request.Form("MURealName"))
' 제목을 읽는다.
strSubject = Trim(Request.Form("MSubject"))
' 단일(') 인용 부호 문제를 없앤다.
strSubject = Replace(strSubject, "'", "''")
' 내용을 읽고, 처리한다. 
strContent = Trim(Request.Form("MContent"))
strContent = Replace(strContent, "'", "''")
' 암호를 읽고, 처리한다. 
strUPassword = Trim(Request.Form("MUPassword"))
strUPassword = Replace(strUPassword, "'", "''")
' 새글의 MID를 결정하기 위하여 현재의 가장 큰 MID를 얻는다.
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
sqlQuery = "SELECT MAX(MID) FROM " & strTable 
Set objRS = objConn.Execute(sqlQuery)
If IsNull(objRS(0)) Then
    intMID = 1
Else
    intMID = objRS(0) + 1
End If
' 또한, 글을 쓴 시각을 구합니다.
strWTime = CStr(Now())  
' 데이타베이스 구조
' (MID,MGrp,MSeq,MLvl,MSubject,MContent,MURealName,MUPassword,MVisited,MWTime) 
' 데이타베이스에 사용자의 입력을 저장한다.
sqlQuery = "INSERT INTO " & strTable & " VALUES ("
sqlQuery = sqlQuery & intMID 
sqlQuery = sqlQuery & ","  & intMID & ", 1, 0"
' 사용자 입력을 저장한다.
sqlQuery = sqlQuery & ",'" & strSubject & "'" 
sqlQuery = sqlQuery & ",'" & strContent & "'" 
sqlQuery = sqlQuery & ",'" & strURealName & "'" 
sqlQuery = sqlQuery & ",'" & strUPassword & "'" 
' 조회수와 글작성 시간을 저장한다.
sqlQuery = sqlQuery & ",1"
sqlQuery = sqlQuery & ",'" & strWTime & "')" 
objConn.Execute sqlQuery
' 사용한 객체들을 반납한다.
objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
' 게시판 목록으로 되돌아 간다.
strURL = "YMList.asp?" & CArg1
Response.Redirect strURL
%>
