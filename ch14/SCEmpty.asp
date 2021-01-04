<%
' Shopping Cart를 찾는다. 
strSCID = Session("sesSCID")
intSCID = CInt(strSCID)
' Shopping Cart에서 해당 레코드들을 삭제한다.
Set objConn = Session("objConn")
sqlQuery = "DELETE FROM SCart WHERE SCID=" & intSCID
objConn.Execute(sqlQuery) 
%> 
<HTML>
<HEAD>
<TITLE> Shopping Cart Empty </TITLE>
</HEAD>
<BODY>
장바구니를 비웠어요.<BR>
<A HREF="SCView.asp">장바구니보기</A>
</BODY>
</HTML>
