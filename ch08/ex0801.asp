<%
strStatus = Session("sesStatus")
%>
<HTML>
<HEAD> 
<TITLE> Membership Application (1st step) </TITLE>
</HEAD>
<BODY>
<I>ȸ���� ���ø����̼�(1st Step)</I> <BR>
<% If strStatus = "" Then %>
    <FONT COLOR="red">ó�� ����ϴ� ����.</FONT>
<% ElseIf strStatus = "1" Then %>
    <FONT COLOR="red">������ �����.</FONT>
<% ElseIf strStatus = "2" Then %>
    <FONT COLOR="red">��ȣ�� Ʋ����.</FONT>
<% End If %>
<FORM METHOD="post" ACTION="ex0802.asp">
<TABLE BORDER=1>
<TR>
    <TD> �� �� </TD>
    <TD> 
	<INPUT TYPE="text" NAME="frmLogin" SIZE=15> </TD>
</TR> <TR>
    <TD> �� ȣ </TD>
    <TD> 
	<INPUT TYPE="password" NAME="frmPassword" SIZE=15> </TD>
</TR> 
</TABLE>
<INPUT TYPE="submit" VALUE="����">
</FORM>
</BODY>
</HTML>
