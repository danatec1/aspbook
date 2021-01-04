<%
' 데이터베이스에 연결할 연결 문자열을 지정한다.
ConnString = "DSN=aspbook;UID=sa;PWD=;"
' 업로드되는 파일들을 저장할 가상 디렉토리를 지정한다.
strVirUDir = "/JFiles"
strPhyUDir = Server.MapPath(strVirUDir)
' 이 컴퓨터의 URL을 지정한다.
strMyURL = "http://localhost"
%>
