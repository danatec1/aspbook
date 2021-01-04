<%
Dim intNumRecs
' 사용자의 입력을 읽는다.
strLogin = Request.Form("frmLogin")
' Command 객체를 생성한다.
Set objComm = Server.CreateObject("ADODB.Command")
' ActiveConnection 속성에 연결 문자열을 지정한다.
ConnString = "DSN=aspbook;UID=sa;PWD=;"
objComm.ActiveConnection = ConnString
' CommandText 속성에 실행할 질의어를 지정한다.
sqlQuery = " DELETE FROM Account WHERE login='" & strLogin & "'"
objComm.CommandText = sqlQuery 
' 질의문을 실행하여 계정을 삭제한다.
objComm.Execute intNumRecs
%>
<HTML>
<HEAD>
<TITLE> User Account Delete </TITLE>
</HEAD>
<BODY>
<% 
If (intNumRecs > 0) Then %>
    회원 계정을 삭제했어요.
<% 
Else %>
    회원 계정이 없어요.
<% 
End If %>
</BODY>
</HTML>
<%
Set objComm = Nothing
%>
