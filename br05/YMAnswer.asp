<!-- 
#include file="YConf.inc" 
-->
<%
' ������� �Է��� �д´�.
strMID = Request.QueryString("MID")
' �̸��� �д´�. 
strURealName = Trim(Request.Form("MURealName"))
' ������ �д´�.
strSubject = Trim(Request.Form("MSubject"))
' ����(') �ο��ȣ ������ ���ش�.
strSubject = Replace(strSubject, "'", "''")
' ������ �а�, ó���Ѵ�. 
strContent = Trim(Request.Form("MContent"))
strContent = Replace(strContent, "'", "''")
' ��ȣ�� �а�, ó���Ѵ�. 
strUPassword = Trim(Request.Form("MUPassword"))
strUPassword = Replace(strUPassword, "'", "''")
'
' ������ MID�� �����ϱ� ���Ͽ� ���� ���� ū MID�� ��´�.
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
sqlQuery = "SELECT MAX(MID) FROM " & strTable 
Set objRS = objConn.Execute(sqlQuery)
If IsNull(objRS(0)) Then
    intNMID = 1
Else
    intNMID = objRS(0) + 1
End If
' ����, ���� �� �ð��� ���մϴ�.
strWTime = CStr(Now())
'
' ���� ���� ���Ͽ� �߰� �۾��� �Ѵ�.
'
' ���� ���� �׷��� �θ�� ���� �׷��̴�.
sqlQuery = "SELECT * FROM " & strTable 
sqlQuery = sqlQuery & " WHERE MID=" & strMID
objRS.Close
Set objRS = objConn.Execute(sqlQuery)
intNMGrp = objRS("MGrp")
' ���� ���� ������ ���� ������ ���Ѵ�. 
intNMSeq = objRS("MSeq") + 1
intNMLvl = objRS("MLvl") + 1
' ���� �޼��� ���Ŀ� �ִ� MSeq �ʵ� ���� ���� �����Ѵ�.
sqlQuery = "UPDATE " & strTable 
sqlQuery = sqlQuery & " SET MSeq=MSeq+1 "
sqlQuery = sqlQuery & " WHERE MGrp =" & intNMGrp 
sqlQuery = sqlQuery & "   AND MSeq >" & objRS("MSeq") 
objConn.Execute(sqlQuery)
'
' ����Ÿ���̽� ����
' (MID,MGrp,MSeq,MLvl,MSubject,MContent,MURealName,MUPassword,MVisited,MWTime) 
' ����Ÿ���̽��� �������� ���õ� �Էµ��� �����Ѵ�.
sqlQuery = "INSERT INTO " & strTable & " VALUES ("
sqlQuery = sqlQuery & intNMID 
sqlQuery = sqlQuery & "," & intNMGrp  
sqlQuery = sqlQuery & "," & intNMSeq  
sqlQuery = sqlQuery & "," & intNMLvl  
' ����� �Է��� �����Ѵ�.
sqlQuery = sqlQuery & ",'" & strSubject & "'" 
sqlQuery = sqlQuery & ",'" & strContent & "'" 
sqlQuery = sqlQuery & ",'" & strURealName & "'" 
sqlQuery = sqlQuery & ",'" & strUPassword & "'" 
' ��ȸ���� ���ۼ� �ð��� �����Ѵ�.
sqlQuery = sqlQuery & ",1"
sqlQuery = sqlQuery & ",'" & strWTime & "')" 
objConn.Execute sqlQuery
' ����� �ڿ��� �ݳ��Ѵ�.
objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
' �Խ��� ������� �ǵ��� ����.
strURL = "YMList.asp?" & CArg1
Response.Redirect strURL
%>
