<% @ TRANSACTION=required %>
<%
' 고객의 입력 정보를 읽는다. 
strUName = Request.Form("UName")
strUAddress = Request.Form("UAddress")
' 새로운 구매 번호를 결정하기 위해서 가장 큰 SOID를 얻는다.
Set objConn = Session("objConn")
sqlQuery = "SELECT MAX(SOID) FROM SOrder"
Set objRS = objConn.Execute(sqlQuery)
If IsNull(objRS(0)) Then
    intSOID = 1
Else
    intSOID = 1 + objRS(0)
End If
' SUser 테이블에 사용자 입력 데이터와 구매 번호를 저장한다.
sqlQuery = "INSERT INTO SUser VALUES (" 
sqlQuery = sqlQuery & "'"  & strUName & "'"  
sqlQuery = sqlQuery & ",'" & strUAddress & "'"
sqlQuery = sqlQuery & ","  & intSOID & ")"
objConn.Execute(sqlQuery) 
' 
' 장바구니의 정보를 구매(Order) 테이블로 복사한다. 
'
' Shopping Cart를 찾는다. 
strSCID = Session("sesSCID")
intSCID = CInt(strSCID)
' 장바구니 데이터들을 읽는다.
sqlQuery = "SELECT * FROM SCart "
sqlQuery = sqlQuery & " WHERE SCID=" & intSCID
Set objRS = objConn.Execute(sqlQuery)
' 구매 테이블로 데이터들을 복사한다.
Do Until objRS.EOF 
    strGCode = objRS("SGCODE")
    intGAmount = objRS("SGAmount")
    sqlQuery = "INSERT INTO SOrder VALUES (" 
    sqlQuery = sqlQuery        & intSOID  
    sqlQuery = sqlQuery & ",'" & strGCode & "'"
    sqlQuery = sqlQuery & ","  & intGAmount & ")"
    objConn.Execute(sqlQuery) 
    objRS.MoveNext
Loop
'
' Shopping Cart를 비운다. 
'
sqlQuery = "DELETE SCart WHERE SCID=" & intSCID
objConn.Execute(sqlQuery) 
%> 
<% 
Sub OnTransactionCommit() %>
    <HTML>
    <HEAD>
    <TITLE> Transaction Success </TITLE>
    </HEAD>
    <BODY>
    구매가 완료되었어요. 상품을 구입해주셔서 감사합니다. <BR>
    <A HREF="SCView.asp">장바구니보기</A>
    </BODY>
    </HTML>
<%
End Sub
Sub OnTransactionAbort() %> 
    <HTML>
    <HEAD>
    <TITLE> Transaction Abort </TITLE>
    </HEAD>
    <BODY>
    구매가 완료되지 않았어요. 다시 시도해주세요. <BR>
    <A HREF="SCView.asp">장바구니보기</A>
    </BODY>
    </HTML>
<%
End Sub 
%>
