<%
' ���� ���� ��ǰ�� ������ �д´�. 
strGCode = Request.Form("GCODE")
strGAmount = Request.Form("GAMOUNT")
intGAmount = CInt(strGAmount)
' Shopping Cart�� ã�´�. 
strSCID = Session("sesSCID")
' Shopping Cart�� ������, Shopping Cart�� ���Ѵ�. 
If strSCID = "" Then
    Application.Lock
    Application("SCID") = 1 + Application("SCID")
    strSCID = Application("SCID")
    Application.UnLock
    Session("sesSCID") = strSCID
End If
intSCID = CInt(strSCID)
' Shopping Cart�� ��ǰ ������ ������ �����Ѵ�.
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
��ٱ��Ͽ� ��ǰ�� �־����. 
</BODY>
</HTML>
