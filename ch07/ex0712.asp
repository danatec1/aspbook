<%
' ��Ű���� �д´�. (�ڵ����� ������ ���۵ȴ�.)
strOldAmount = Request.Cookies("cooAmount")
intOldAmount = CInt(strOldAmount)
' ���ο� �Աݾ��� �д´�.
strMoney = Request.Form("varMoney")
intMoney = CInt(strMoney)
' ���ο� �ܾ��� ����Ѵ�.
'   �ܾ� = �Աݾ� + ������ �ܾ�
intAmount = intMoney + intOldAmount
' ��Ű���� �缳���Ѵ�.(Ŭ���̾�Ʈ�� ��Ű���� �����Ѵ�.)
Response.Cookies("cooAmount") = intAmount
%>
<HTML>
<HEAD> 
<TITLE> Cookie Bank </TITLE>
</HEAD>
<BODY>
<I>���� ����</I> <BR>
���ο� �Ա� : <% = intMoney %> �� <BR>
������ �ܾ� : <% = intOldAmount %> �� <BR>
�ܾ� : <% = intAmount %> �� <BR>
<FORM METHOD="post" ACTION="ex0712.asp">
�ݾ� : <INPUT TYPE="text" NAME="varMoney" SIZE=10> �� <BR>
<INPUT TYPE="submit" VALUE="�߰��Ա�"> 
</FORM>
</BODY>
</HTML>
