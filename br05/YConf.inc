<%
' 데이터베이스에 연결할 연결 문자열을 지정한다.
ConnString = "DSN=ybbs;UID=sa;PWD=;"
'
' 모든 웹 페이지에 전송되는 공통 입력 인수들을 읽는다.
'
' Table : 테이블 이름(게시판)을 읽는다. (다중 게시판)
strTable=Request.QueryString("Table")
CArg1  = "Table=" & strTable
' Mode : 게시물을 보여주는 방법을 정한다. (순서형과 응답형)
strMode=Request.QueryString("Mode")
If (strMode = "") Then
    strMode = "qa"
End If
CArg2  = CArg1 & "&Mode=" & strMode   
' Page : 디스플레이할 현재 페이지 번호를 읽는다.   
strCPage = Request.QueryString("Page")
If (strCPage = "") Then
    strCPage = "1"
End If 
CArg3  = CArg2 & "&Page=" & strCPage  
intCPage = CInt(strCPage)
%>
