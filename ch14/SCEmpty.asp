<%
' Shopping Cart�� ã�´�. 
strSCID = Session("sesSCID")
intSCID = CInt(strSCID)
' Shopping Cart���� �ش� ���ڵ���� �����Ѵ�.
Set objConn = Session("objConn")
sqlQuery = "DELETE FROM SCart WHERE SCID=" & intSCID
objConn.Execute(sqlQuery) 
%> 
<HTML>
<HEAD>
<TITLE> Shopping Cart Empty </TITLE>
</HEAD>
<BODY>
��ٱ��ϸ� ������.<BR>
<A HREF="SCView.asp">��ٱ��Ϻ���</A>
</BODY>
</HTML>
