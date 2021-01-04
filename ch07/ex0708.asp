<% 
strCard = Request.Form("card")
%>
<HTML>
<HEAD> 
<TITLE> Card Sending : Content Form </TITLE>
</HEAD>
<BODY>
<I>카드 보내기</I> <BR>
전자메일 주소와 원하는 내용을 쓰세요.(2nd Step)
<FORM METHOD="post" ACTION="ex0709.asp">
<TABLE BORDER=1>
<!-- 첫번째 입력된 데이터 값을 두번째 입력 양식에 끼워 넣는다. -->
<TR>
    <TD COLSPAN=2> 
	<INPUT TYPE="hidden" NAME="card" VALUE="<%=strCard%>"> </TD>
</TR> <TR>
    <TD> 전자메일 주소 </TD>
    <TD> 
	<INPUT TYPE="text" NAME="email" SIZE=30> </TD>
</TR> <TR>
    <TD> 하고픈 이야기 </TD>
    <TD> 
	<TEXTAREA ROWS=3 COLS=30 NAME="content"></TEXTAREA> </TD>
</TR> 
</TABLE>
<INPUT TYPE="submit" VALUE="입력완료">
</FORM>
</BODY>
</HTML>
