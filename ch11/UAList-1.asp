<%
' Recordset 객체를 생성한 후에
' 데이터베이스에 연결하여 레코드셋을 읽는다.
Set objRS = Server.CreateObject("ADODB.Recordset")
ConnString = "DSN=aspbook;UID=sa;PWD=;"
sqlQuery = "SELECT * FROM account" 
objRS.Open sqlQuery, ConnString
' HTML 문서의 앞 부분을 출력한다.
Response.Write "<HTML>" 
Response.Write "<HEAD>"
Response.Write "<TITLE> User Account List </TITLE>"
Response.Write "</HEAD>"
Response.Write "<BODY>"
Response.Write "<I>회원 목록</I>"
' 레코드셋에 레코드들이 있는가 ?
If (objRS.BOF and objRS.EOF) Then
    Response.Write "리턴 결과가 없어요."
Else
' 레코드셋에 레코드가 있으면,
    FCount = objRS.Fields.Count - 1
    Response.Write "<TABLE BORDER=1>"
    Response.Write "<TR>"
'   필드 이름을 출력한다.
    For i=0 To FCount
	Response.Write "<TH> " & objRS(i).Name & " </TH>"
    Next
    Response.Write "</TR>"
'   레코드들을 하나씩 출력한다.
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
' HTML 문서의 뒷 부분을 출력한다.
Response.Write "</BODY>"
Response.Write "</HTML>"
' CursorType을 테스트하기 위한 것입니다.
' objRS.MovePrevious
' Recordset 객체를 닫고 파기한다.
objRS.Close
Set objRS = Nothing
%>
