<%
strStatus = Session("sesStatus")
%>
<HTML>
<HEAD> 
<TITLE> Membership Application (1st step) </TITLE>
</HEAD>
<BODY>
<I>회원제 어플리케이션(1st Step)</I> <BR>
<% If strStatus = "" Then %>
    <FONT COLOR="red">처음 사용하는 군요.</FONT>
<% ElseIf strStatus = "1" Then %>
    <FONT COLOR="red">계정이 없어요.</FONT>
<% ElseIf strStatus = "2" Then %>
    <FONT COLOR="red">암호가 틀려요.</FONT>
<% End If %>
<FORM METHOD="post" ACTION="ex0802.asp">
<TABLE BORDER=1>
<TR>
    <TD> 계 정 </TD>
    <TD> 
	<INPUT TYPE="text" NAME="frmLogin" SIZE=15> </TD>
</TR> <TR>
    <TD> 암 호 </TD>
    <TD> 
	<INPUT TYPE="password" NAME="frmPassword" SIZE=15> </TD>
</TR> 
</TABLE>
<INPUT TYPE="submit" VALUE="입장">
</FORM>
</BODY>
</HTML>
