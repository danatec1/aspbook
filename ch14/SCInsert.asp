<%
' 고객의 구입 상품의 정보를 읽는다. 
strGCode = Request.Form("GCODE")
strGAmount = Request.Form("GAMOUNT")
intGAmount = CInt(strGAmount)
' Shopping Cart를 찾는다. 
strSCID = Session("sesSCID")
' Shopping Cart가 없으면, Shopping Cart를 구한다. 
If strSCID = "" Then
    Application.Lock
    Application("SCID") = 1 + Application("SCID")
    strSCID = Application("SCID")
    Application.UnLock
    Session("sesSCID") = strSCID
End If
intSCID = CInt(strSCID)
' Shopping Cart에 상품 종류와 갯수를 저장한다.
Set objConn = Session("objConn")
sqlQuery = "INSERT INTO SCart VALUES (" 
sqlQuery = sqlQuery        & intSCID  
sqlQuery = sqlQuery & ",'" & strGCode & "'"
sqlQuery = sqlQuery & ","  & intGAmount & ")"
objConn.Execute(sqlQuery) 
%> 
<HTML>
<HEAD>
<TITLE> Goods Buy </TITLE>
</HEAD>
<BODY>
장바구니에 상품을 넣었어요. 
</BODY>
</HTML>
