<%
' �����ͺ��̽��� ������ ���� ���ڿ��� �����Ѵ�.
ConnString = "DSN=ybbs;UID=sa;PWD=;"
'
' ��� �� �������� ���۵Ǵ� ���� �Է� �μ����� �д´�.
'
' Table : ���̺� �̸�(�Խ���)�� �д´�. (���� �Խ���)
strTable=Request.QueryString("Table")
CArg1  = "Table=" & strTable
' Mode : �Խù��� �����ִ� ����� ���Ѵ�. (�������� ������)
strMode=Request.QueryString("Mode")
If (strMode = "") Then
    strMode = "qa"
End If
CArg2  = CArg1 & "&Mode=" & strMode   
' Page : ���÷����� ���� ������ ��ȣ�� �д´�.   
strCPage = Request.QueryString("Page")
If (strCPage = "") Then
    strCPage = "1"
End If 
CArg3  = CArg2 & "&Page=" & strCPage  
intCPage = CInt(strCPage)
%>