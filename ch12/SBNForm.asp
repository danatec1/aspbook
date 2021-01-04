<HTML>
<HEAD>
<TITLE> Form for New Message </TITLE>
</HEAD>
<BODY>
<I>새글쓰기</I>
<TABLE>
<FORM METHOD="post" ACTION="SBNew.asp">
<TR>
    <TD> 이름 </TD>
    <TD> <INPUT TYPE="text" NAME="frmUName" maxlength="10"> </TD>
</TR> <TR>
    <TD> 제목 </TD>
    <TD> <INPUT TYPE="text" NAME="frmSubject" SIZE=55> </TD>
</TR> <TR>
    <TD> 내용 </TD>
    <TD> <TEXTAREA COLS="55" ROWS="5" NAME="frmContent"></textarea> </TD>
</TR> <TR>       
    <TD COLSPAN=2 ALIGN="center">
	<INPUT TYPE="submit" VALUE="등록">
	<INPUT TYPE="reset"  VALUE="재작성">
    </TD>
</FORM>
</TABLE>
</BODY>
</HTML>
