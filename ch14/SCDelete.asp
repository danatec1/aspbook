<%
' Shopping Cart를 찾는다. 
strSCID = Session("sesSCID")
intSCID = CInt(strSCID)
' 프로그램의 입력(상품 코드)을 읽는다.
strGCODE = Request.QueryString("GCODE")
' Shopping Cart에서 해당 레코드를 삭제한다.
Set objConn = Session("objConn")
sqlQuery = "DELETE SCart WHERE SCID=" & intSCID
sqlQuery = sqlQuery & "  and SGCODE='" & strGCODE & "'" 
objConn.Execute(sqlQuery) 
' 장바구니 보기로 되돌아간다.
Response.Redirect "SCView.asp"
%> 
