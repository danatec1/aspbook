<HTML>
<HEAD>
<TITLE> Guest Book : New Form </TITLE>
</HEAD>
<BODY>
글쓰기
<FORM METHOD="post" ACTION="GBNew.asp">
<TABLE>
<TR>
    <TD> 이름 </TD>
    <TD> <INPUT TYPE="text" NAME="frmName" maxlength="10"> </TD>
</TR> <TR>
    <TD> 내용 </TD>
    <TD> <TEXTAREA COLS="50" ROWS="5" NAME="frmContent"></textarea> </TD>
</TR> <TR>       
    <TD COLSPAN=2 ALIGN="center">
	<INPUT TYPE="submit" VALUE="등록">
	<INPUT TYPE="reset"  VALUE="재작성">
    </TD>
</TR>
</TABLE>
</FORM>
</BODY>
</HTML>
