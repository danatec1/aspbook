<!-- 
#include file="YConf.inc" 
-->
<%
' 사용자의 입력을 읽고 인수 목록도 만든다.
strMID = Request.QueryString("MID")
' 해당 게시물을 읽는다.
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
sqlQuery = "SELECT * FROM " & strTable 
sqlQuery = sqlQuery & " WHERE MID =" & strMID
Set objRS = objConn.Execute(SQLQuery)
' 암호를 읽는다.
strUPassword = Request.Form("MUPassword")
If ( objRS("MUPassword") = strUPassword) then
'   해당 게시물을 삭제한다.
    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.Open ConnString
    sqlQuery = "DELETE " & strTable 
    sqlQuery = sqlQuery & " WHERE MID =" & strMID
    objConn.Execute(sqlQuery)
    objConn.Close
    Set objConn = Nothing
    strURL = "YMList.asp?" & CArg3
    Response.Redirect strURL
Else
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
