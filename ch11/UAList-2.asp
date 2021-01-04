<%
' 라인의 끝을 나타내는 문자열을 정의한다.
strCRNL = chr(13) & chr(10)
' Recordset 객체를 생성한 후에
' 데이터베이스에 연결하여 레코드셋을 읽는다.
Set objRS = Server.CreateObject("ADODB.Recordset")
ConnString = "DSN=aspbook;UID=sa;PWD=;"
sqlQuery = "SELECT * FROM account" 
objRS.Open sqlQuery, ConnString
' HTML 문서의 앞 부분을 출력한다.
Response.Write "<HTML>" & strCRNL
Response.Write "<HEAD>" & strCRNL
Response.Write "<TITLE> User Account List </TITLE>" & strCRNL
Response.Write "</HEAD>" & strCRNL
Response.Write "<BODY>" & strCRNL
Response.Write "<I>회원 목록</I>" & strCRNL
' 레코드셋에 레코드들이 있는가 ?
If (objRS.BOF and objRS.EOF) Then
    Response.Write "리턴 결과가 없어요." & strCRNL
Else
' 레코드셋에 레코드가 있으면,
    FCount = objRS.Fields.Count - 1
    Response.Write "<TABLE BORDER=1>" & strCRNL
    Response.Write "<TR>" & strCRNL
'   필드 이름을 출력한다.
    For i=0 To FCount
	Response.Write "<TH> " & objRS(i).Name & " </TH>" & strCRNL
    Next
    Response.Write "</TR>" & strCRNL
'   레코드들을 하나씩 출력한다.
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
' HTML 문서의 뒷 부분을 출력한다.
Response.Write "</BODY>" & strCRNL
Response.Write "</HTML>" & strCRNL
' CursorType을 테스트하기 위한 것입니다.
' objRS.MovePrevious
' Recordset 객체를 닫고 파기한다.
objRS.Close
Set objRS = Nothing
%>
