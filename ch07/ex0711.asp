<%
strMoney = Request.Form("varMoney")
' ��Ű���� �����Ѵ�.(Ŭ���̾�Ʈ�� �����Ѵ�.)
Response.Cookies("cooAmount") = strMoney
%>
<HTML>
<HEAD> 
<TITLE> Cookie Bank </TITLE>
</HEAD>
<BODY>
<I>���� ����</I> <BR>
�ܾ� : <% = strMoney %> �� 
<FORM METHOD="post" ACTION="ex0712.asp">
�ݾ� : <INPUT TYPE="text" NAME="varMoney" SIZE=10> �� <BR>
<INPUT TYPE="submit" VALUE="�Ա�">
</FORM>
</BODY>
</HTML>
