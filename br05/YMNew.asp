<!-- 
#include file="YConf.inc" 
-->
<%
' �̸��� �д´�. 
strURealName = Trim(Request.Form("MURealName"))
' ������ �д´�.
strSubject = Trim(Request.Form("MSubject"))
' ����(') �ο� ��ȣ ������ ���ش�.
strSubject = Replace(strSubject, "'", "''")
' ������ �а�, ó���Ѵ�. 
strContent = Trim(Request.Form("MContent"))
strContent = Replace(strContent, "'", "''")
' ��ȣ�� �а�, ó���Ѵ�. 
strUPassword = Trim(Request.Form("MUPassword"))
strUPassword = Replace(strUPassword, "'", "''")
' ������ MID�� �����ϱ� ���Ͽ� ������ ���� ū MID�� ��´�.
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
sqlQuery = "SELECT MAX(MID) FROM " & strTable 
Set objRS = objConn.Execute(sqlQuery)
If IsNull(objRS(0)) Then
    intMID = 1
Else
    intMID = objRS(0) + 1
End If
' ����, ���� �� �ð��� ���մϴ�.
strWTime = CStr(Now())  
' ����Ÿ���̽� ����
' (MID,MGrp,MSeq,MLvl,MSubject,MContent,MURealName,MUPassword,MVisited,MWTime) 
' ����Ÿ���̽��� ������� �Է��� �����Ѵ�.
sqlQuery = "INSERT INTO " & strTable & " VALUES ("
sqlQuery = sqlQuery & intMID 
sqlQuery = sqlQuery & ","  & intMID & ", 1, 0"
' ����� �Է��� �����Ѵ�.
sqlQuery = sqlQuery & ",'" & strSubject & "'" 
sqlQuery = sqlQuery & ",'" & strContent & "'" 
sqlQuery = sqlQuery & ",'" & strURealName & "'" 
sqlQuery = sqlQuery & ",'" & strUPassword & "'" 
' ��ȸ���� ���ۼ� �ð��� �����Ѵ�.
sqlQuery = sqlQuery & ",1"
sqlQuery = sqlQuery & ",'" & strWTime & "')" 
objConn.Execute sqlQuery
' ����� ��ü���� �ݳ��Ѵ�.
objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
' �Խ��� ������� �ǵ��� ����.
strURL = "YMList.asp?" & CArg1
Response.Redirect strURL
%>
