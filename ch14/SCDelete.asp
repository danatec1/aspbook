<%
' Shopping Cart�� ã�´�. 
strSCID = Session("sesSCID")
intSCID = CInt(strSCID)
' ���α׷��� �Է�(��ǰ �ڵ�)�� �д´�.
strGCODE = Request.QueryString("GCODE")
' Shopping Cart���� �ش� ���ڵ带 �����Ѵ�.
Set objConn = Session("objConn")
sqlQuery = "DELETE SCart WHERE SCID=" & intSCID
sqlQuery = sqlQuery & "  and SGCODE='" & strGCODE & "'" 
objConn.Execute(sqlQuery) 
' ��ٱ��� ����� �ǵ��ư���.
Response.Redirect "SCView.asp"
%> 
