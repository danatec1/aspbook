<%
' ���ο� �Աݾ��� �д´�.
strMoney = Request.Form("varMoney")
intMoney = CInt(strMoney)
%>
<HTML>
<HEAD> 
<TITLE> Example : Blackhole Bank (Result) </TITLE>
</HEAD>
<BODY>
<I>���� ����</I> <BR>
���ο� �Ա� : <% = intMoney %> �� <BR>
������ �ܾ� : <% = intAmount %> �� <BR>
<BR>
<%
' ���ο� �ܾ��� ����Ѵ�.
'   �ܾ� = �Աݾ� + �ܾ�
intAmount = intMoney + intAmount
%>
�ܾ� : <% = intAmount %> �� <BR>
</BODY>
</HTML>
