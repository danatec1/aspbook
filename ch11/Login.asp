<%
strStatus = Request.QueryString("Status")
%>
<HTML>
<HEAD> 
<TITLE>Member Login Form</TITLE>
</HEAD>
<BODY>
<A HREF="UANForm.asp">ȸ�� ����</A> <BR>
<TABLE BORDER=1>
<FORM METHOD="post" ACTION="UACheck.asp">
<TR>
    <TD> �� �� </TD>
    <TD> 
	<INPUT TYPE="text" NAME="frmLogin" SIZE=15> </TD>
</TR> <TR>
    <TD> �� ȣ </TD>
    <TD> 
	<INPUT TYPE="password" NAME="frmPassword" SIZE=15> </TD>
</TR> 
<TR>
    <TD COLSPAN=2>
    <INPUT TYPE="submit" VALUE="����">
<%
If strStatus = "1" Then %>
    <FONT COLOR="red">������ �����.</FONT>
<% 
ElseIf strStatus = "2" Then %>
    <FONT COLOR="red">��ȣ�� Ʋ����.</FONT>
<% 
Else %>
    <FONT COLOR="red">ȸ���� �����ϼ���.</FONT>
<% 
End If %>
    </TD>
</TR>
</FORM>
</TABLE>
</BODY>
</HTML>
