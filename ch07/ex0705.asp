<%
strMoney = Request.Form("varMoney")
intAmount = CInt(strMoney)
%>
<HTML>
<HEAD> 
<TITLE> Example : Blackhole Bank </TITLE>
</HEAD>
<BODY>
<I>���� ����</I> <BR>
�ܾ� : <% = intAmount %> �� 
<FORM METHOD="post" ACTION="ex0706.asp">
�ݾ� : <INPUT TYPE="text" NAME="varMoney" SIZE=10> �� <BR>
<INPUT TYPE="submit" VALUE="�Ա�">
</FORM>
</BODY>
</HTML>
