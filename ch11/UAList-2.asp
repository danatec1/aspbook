<%
' ������ ���� ��Ÿ���� ���ڿ��� �����Ѵ�.
strCRNL = chr(13) & chr(10)
' Recordset ��ü�� ������ �Ŀ�
' �����ͺ��̽��� �����Ͽ� ���ڵ���� �д´�.
Set objRS = Server.CreateObject("ADODB.Recordset")
ConnString = "DSN=aspbook;UID=sa;PWD=;"
sqlQuery = "SELECT * FROM account" 
objRS.Open sqlQuery, ConnString
' HTML ������ �� �κ��� ����Ѵ�.
Response.Write "<HTML>" & strCRNL
Response.Write "<HEAD>" & strCRNL
Response.Write "<TITLE> User Account List </TITLE>" & strCRNL
Response.Write "</HEAD>" & strCRNL
Response.Write "<BODY>" & strCRNL
Response.Write "<I>ȸ�� ���</I>" & strCRNL
' ���ڵ�¿� ���ڵ���� �ִ°� ?
If (objRS.BOF and objRS.EOF) Then
    Response.Write "���� ����� �����." & strCRNL
Else
' ���ڵ�¿� ���ڵ尡 ������,
    FCount = objRS.Fields.Count - 1
    Response.Write "<TABLE BORDER=1>" & strCRNL
    Response.Write "<TR>" & strCRNL
'   �ʵ� �̸��� ����Ѵ�.
    For i=0 To FCount
	Response.Write "<TH> " & objRS(i).Name & " </TH>" & strCRNL
    Next
    Response.Write "</TR>" & strCRNL
'   ���ڵ���� �ϳ��� ����Ѵ�.
    While (Not objRS.EOF)
	Response.Write "<TR>" & strCRNL
        For i=0 To FCount
	    Response.Write "<TD> " & objRS(i) & " </TD>" & strCRNL
        Next
	Response.Write "</TR>" & strCRNL
	objRS.MoveNext
    Wend
    Response.Write "</TABLE>" & strCRNL
End If
' HTML ������ �� �κ��� ����Ѵ�.
Response.Write "</BODY>" & strCRNL
Response.Write "</HTML>" & strCRNL
' CursorType�� �׽�Ʈ�ϱ� ���� ���Դϴ�.
' objRS.MovePrevious
' Recordset ��ü�� �ݰ� �ı��Ѵ�.
objRS.Close
Set objRS = Nothing
%>
