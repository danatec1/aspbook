<%
' Shopping Cart�� ã�´�. 
strSCID = Session("sesSCID")
intSCID = CInt(strSCID)
' ���� �Է� �����͵��� �д´�.
strGCODE = Request.Form("GCODE")
strGAmount = Request.Form("GAMOUNT")
intGAmount = CInt(strGAmount)
' �����ͺ��̽��� �����͵��� �����Ѵ�.
Set objConn = Session("objConn")
sqlQuery = "UPDATE SCart SET " 
sqlQuery = sqlQuery & " SGAMOUNT = " & intGAmount  
sqlQuery = sqlQuery & " WHERE SCID=" & intSCID
sqlQuery = sqlQuery & " and SGCODE='" & strGCODE & "'"
objConn.Execute(sqlQuery) 
' ��ٱ��� ����� �ǵ��ư���.
response.redirect "SCView.asp"
%> 
