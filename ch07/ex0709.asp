<%
' Hidden Field���� �д´�.
strCard = Request.Form("card")
' ���� �Էµ� �����͵��� �д´�.
strEmail = Request.Form("email")
strContent = Request.Form("content")
%>
<HTML>
<HEAD> 
<TITLE> Card Sending : Card Sending </TITLE>
</HEAD>
<BODY>
<I>ī�� ������</I> <BR>
������ ���� �����͸� ó���Ͽ� ���ڸ��Ϸ� ������.(3rd Step) <BR>
(������ ���ڸ����� ������ ����� �η��� ����)  
<TABLE BORDER=1>
<TR>
    <TD> �׸� </TD>
    <TD> <% = strCard %> </TD>
</TR> <TR>
    <TD> �ּ� </TD>
    <TD> <% = strEMail %> </TD>
</TR> <TR>
    <TD> ���� </TD>
    <TD> <% = strContent %> </TD>
</TR>
</BODY>
</HTML>
