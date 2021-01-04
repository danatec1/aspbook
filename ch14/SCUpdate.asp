<%
' Shopping Cart를 찾는다. 
strSCID = Session("sesSCID")
intSCID = CInt(strSCID)
' 고객의 입력 데이터들을 읽는다.
strGCODE = Request.Form("GCODE")
strGAmount = Request.Form("GAMOUNT")
intGAmount = CInt(strGAmount)
' 데이터베이스에 데이터들을 수정한다.
Set objConn = Session("objConn")
sqlQuery = "UPDATE SCart SET " 
sqlQuery = sqlQuery & " SGAMOUNT = " & intGAmount  
sqlQuery = sqlQuery & " WHERE SCID=" & intSCID
sqlQuery = sqlQuery & " and SGCODE='" & strGCODE & "'"
objConn.Execute(sqlQuery) 
' 장바구니 보기로 되돌아간다.
response.redirect "SCView.asp"
%> 
