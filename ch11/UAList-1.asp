<%
' Recordset ��ü�� ������ �Ŀ�
' �����ͺ��̽��� �����Ͽ� ���ڵ���� �д´�.
Set objRS = Server.CreateObject("ADODB.Recordset")
ConnString = "DSN=aspbook;UID=sa;PWD=;"
sqlQuery = "SELECT * FROM account" 
objRS.Open sqlQuery, ConnString
' HTML ������ �� �κ��� ����Ѵ�.
Response.Write "<HTML>" 
Response.Write "<HEAD>"
Response.Write "<TITLE> User Account List </TITLE>"
Response.Write "</HEAD>"
Response.Write "<BODY>"
Response.Write "<I>ȸ�� ���</I>"
' ���ڵ�¿� ���ڵ���� �ִ°� ?
If (objRS.BOF and objRS.EOF) Then
    Response.Write "���� ����� �����."
Else
' ���ڵ�¿� ���ڵ尡 ������,
    FCount = objRS.Fields.Count - 1
    Response.Write "<TABLE BORDER=1>"
    Response.Write "<TR>"
'   �ʵ� �̸��� ����Ѵ�.
    For i=0 To FCount
	Response.Write "<TH> " & objRS(i).Name & " </TH>"
    Next
    Response.Write "</TR>"
'   ���ڵ���� �ϳ��� ����Ѵ�.
    While (Not objRS.EOF)
	Response.Write "<TR>"
        For i=0 To FCount
	    Response.Write "<TD> " & objRS(i) & " </TD>"
        Next
	Response.Write "</TR>"
	objRS.MoveNext
    Wend
    Response.Write "</TABLE>"
End If
' HTML ������ �� �κ��� ����Ѵ�.
Response.Write "</BODY>"
Response.Write "</HTML>"
' CursorType�� �׽�Ʈ�ϱ� ���� ���Դϴ�.
' objRS.MovePrevious
' Recordset ��ü�� �ݰ� �ı��Ѵ�.
objRS.Close
Set objRS = Nothing
%>
