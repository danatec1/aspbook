<%
strStatus = Request.QueryString("Status")
%>
<HTML>
<HEAD> 
<TITLE>Member Login Form</TITLE>
</HEAD>
<BODY>
<A HREF="UANForm.asp">회원 가입</A> <BR>
<TABLE BORDER=1>
<FORM METHOD="post" ACTION="UACheck.asp">
<TR>
    <TD> 계 정 </TD>
    <TD> 
	<INPUT TYPE="text" NAME="frmLogin" SIZE=15> </TD>
</TR> <TR>
    <TD> 암 호 </TD>
    <TD> 
	<INPUT TYPE="password" NAME="frmPassword" SIZE=15> </TD>
</TR> 
<TR>
    <TD COLSPAN=2>
    <INPUT TYPE="submit" VALUE="입장">
<%
If strStatus = "1" Then %>
    <FONT COLOR="red">계정이 없어요.</FONT>
<% 
ElseIf strStatus = "2" Then %>
    <FONT COLOR="red">암호가 틀려요.</FONT>
<% 
Else %>
    <FONT COLOR="red">회원에 가입하세요.</FONT>
<% 
End If %>
    </TD>
</TR>
</FORM>
</TABLE>
</BODY>
</HTML>
