<%
strTable = "Users"
strName = "OGSon"
SQLQuery = "SELECT * FROM " & strTable
SQLQuery = SQLQuery & " WHERE UName=" + strName

strX = "10"
strY = "20"
intX = 10
Sum1 = strX + strY
Sum2 = intX + strY
%>
<HTML>
<HEAD>
<TITLE> Example : Concatenation Operators </TITLE>
</HEAD>
<BODY>
���� �����ڵ�
<HR>
���� ��� ��� : <% = SQLQuery %> <BR>
�߸���(����ġ ����) ��� (���ڿ� + ���ڿ�) <BR> 
SUM1 : <% = Sum1 %> <BR>
�������� ��� (���� + ���ڿ�) <BR>
SUM2 : <% = Sum2 %> 
</BODY>
</HTML>
