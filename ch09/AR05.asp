<%
' �α� ���Ͽ� ����� �ð��� URL�� ���Ѵ�.
strNow = Now()
adHomePageURL = Request.QueryString("url")
strLog = strNow & "  " & adHomePageURL
' �α� ���Ͽ� �ð��� URL�� ����Ѵ�.
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
strPath = Server.MapPath("/booksrc/ch09")
strPFile = strPath & "/" & "ARLog.txt"
' ���Ͽ� ������ ÷��(8)�ϱ� ���ؼ� ����,
' ������ �������� ������ ������ ����(True)�Ѵ�.
Set objTS = objFSO.OpenTextFile(strPFile, 8, True)
objTS.WriteLine(strLog)
objTS.Close
Set objTS = Nothing
' �������� Ȩ�������� �̵��Ѵ�.
Response.Redirect adHomePageURL
%>
