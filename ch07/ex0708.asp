<% 
strCard = Request.Form("card")
%>
<HTML>
<HEAD> 
<TITLE> Card Sending : Content Form </TITLE>
</HEAD>
<BODY>
<I>ī�� ������</I> <BR>
���ڸ��� �ּҿ� ���ϴ� ������ ������.(2nd Step)
<FORM METHOD="post" ACTION="ex0709.asp">
<TABLE BORDER=1>
<!-- ù��° �Էµ� ������ ���� �ι�° �Է� ��Ŀ� ���� �ִ´�. -->
<TR>
    <TD COLSPAN=2> 
	<INPUT TYPE="hidden" NAME="card" VALUE="<%=strCard%>"> </TD>
</TR> <TR>
    <TD> ���ڸ��� �ּ� </TD>
    <TD> 
	<INPUT TYPE="text" NAME="email" SIZE=30> </TD>
</TR> <TR>
    <TD> �ϰ��� �̾߱� </TD>
    <TD> 
	<TEXTAREA ROWS=3 COLS=30 NAME="content"></TEXTAREA> </TD>
</TR> 
</TABLE>
<INPUT TYPE="submit" VALUE="�Է¿Ϸ�">
</FORM>
</BODY>
</HTML>
