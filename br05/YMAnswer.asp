<!-- 
#include file="YConf.inc" 
-->
<%
' 사용자의 입력을 읽는다.
strMID = Request.QueryString("MID")
' 이름을 읽는다. 
strURealName = Trim(Request.Form("MURealName"))
' 제목을 읽는다.
strSubject = Trim(Request.Form("MSubject"))
' 단일(') 인용부호 문제를 없앤다.
strSubject = Replace(strSubject, "'", "''")
' 내용을 읽고, 처리한다. 
strContent = Trim(Request.Form("MContent"))
strContent = Replace(strContent, "'", "''")
' 암호를 읽고, 처리한다. 
strUPassword = Trim(Request.Form("MUPassword"))
strUPassword = Replace(strUPassword, "'", "''")
'
' 새글의 MID를 결정하기 위하여 현재 가장 큰 MID를 얻는다.
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
sqlQuery = "SELECT MAX(MID) FROM " & strTable 
Set objRS = objConn.Execute(sqlQuery)
If IsNull(objRS(0)) Then
    intNMID = 1
Else
    intNMID = objRS(0) + 1
End If
' 또한, 글을 쓴 시각을 구합니다.
strWTime = CStr(Now())
'
' 응답 글을 위하여 추가 작업을 한다.
'
' 응답 글의 그룹은 부모와 같은 그룹이다.
sqlQuery = "SELECT * FROM " & strTable 
sqlQuery = sqlQuery & " WHERE MID=" & strMID
objRS.Close
Set objRS = objConn.Execute(sqlQuery)
intNMGrp = objRS("MGrp")
' 응답 글의 순서와 응답 레벨을 정한다. 
intNMSeq = objRS("MSeq") + 1
intNMLvl = objRS("MLvl") + 1
' 현재 메세지 이후에 있는 MSeq 필드 값을 전부 갱신한다.
sqlQuery = "UPDATE " & strTable 
sqlQuery = sqlQuery & " SET MSeq=MSeq+1 "
sqlQuery = sqlQuery & " WHERE MGrp =" & intNMGrp 
sqlQuery = sqlQuery & "   AND MSeq >" & objRS("MSeq") 
objConn.Execute(sqlQuery)
'
' 데이타베이스 구조
' (MID,MGrp,MSeq,MLvl,MSubject,MContent,MURealName,MUPassword,MVisited,MWTime) 
' 데이타베이스에 응답형과 관련된 입력들을 저장한다.
sqlQuery = "INSERT INTO " & strTable & " VALUES ("
sqlQuery = sqlQuery & intNMID 
sqlQuery = sqlQuery & "," & intNMGrp  
sqlQuery = sqlQuery & "," & intNMSeq  
sqlQuery = sqlQuery & "," & intNMLvl  
' 사용자 입력을 저장한다.
sqlQuery = sqlQuery & ",'" & strSubject & "'" 
sqlQuery = sqlQuery & ",'" & strContent & "'" 
sqlQuery = sqlQuery & ",'" & strURealName & "'" 
sqlQuery = sqlQuery & ",'" & strUPassword & "'" 
' 조회수와 글작성 시간을 저장한다.
sqlQuery = sqlQuery & ",1"
sqlQuery = sqlQuery & ",'" & strWTime & "')" 
objConn.Execute sqlQuery
' 사용한 자원을 반납한다.
objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
' 게시판 목록으로 되돌아 간다.
strURL = "YMList.asp?" & CArg1
Response.Redirect strURL
%>
