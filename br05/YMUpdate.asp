<!-- 
#include file="YConf.inc" 
-->
<%
' 사용자의 입력을 읽고 인수 목록도 만든다.
strMID = Request.QueryString("MID")
CArg4 = CArg3 & "&MID=" & strMID
' 해당 게시물을 읽는다.
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
sqlQuery = "SELECT * FROM " & strTable 
sqlQuery = sqlQuery & " WHERE MID =" & strMID
Set objRS = objConn.Execute(SQLQuery)
'
strUPassword = Request.Form("MUPassword")
If ( objRS("MUPassword") = strUPassword) then
' 사용자들의 데이터들을 읽는다.
    strURealName = Request.Form("MURealName")
    strSubject = Request.Form("MSubject")
    strSubject = Replace(strSubject, "'", "''")
    strContent = Request.Form("MContent")
    strContent = Replace(strContent, "'", "''")
' 해당 레코드를 수정한다.
    sqlQuery = "UPDATE " & strTable 
    sqlQuery = sqlQuery & " SET MURealName='" & strURealName & "'," 
    sqlQuery = sqlQuery & "     MSubject='" & strSubject & "'," 
    sqlQuery = sqlQuery & "     MContent='" & strContent & "'" 
    sqlQUery = sqlQuery & " WHERE MID =" & strMID
    objConn.Execute(SQLQuery)
    objRS.Close
    Set objRS = Nothing
    objConn.Close
    Set objConn = Nothing
    strURL = "YMUForm.asp?" & CArg4
    Response.Redirect strURL
Else
    objRS.Close
    Set objRS=Nothing
    objConn.Close
    Set objConn=Nothing
%>
    <HTML>
    <HEAD>
    <TITLE> Password Error </TITLE>
    </HEAD>
    <BODY>
    암호가 틀립니다. <BR>
    <A HREF="javascript:history.back()">되돌아가기</A>
    </BODY>
    </HTML>
<%
End If
%>
