<%
' �Է� �����͵��� �д´�.
strX = Request.Form.Item("varX")
strY = Request.Form("varY")
' �Է� �����͵��� ���� Ÿ������ ��ȯ�ϰ� ���ϱ� �Ѵ�.
intX = CInt(strX)
intY = CInt(strY)
intSum = intX + intY
%>
<HTML>
<HEAD> 
<TITLE> Example : Form Collection(Result) </TITLE>
</HEAD>
<BODY>
<I>Form �÷���(������ ó��)</I><BR>
�Է� ������ �� : 
X = <% = strX %>, Y = <% = strY %> <BR>
�� ���� �� : <% = intSum %>
</BODY>
</HTML>
