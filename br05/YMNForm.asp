<!-- 
#include file="YConf.inc" 
-->
<HTML>
<HEAD>
<TITLE> Message New Form </TITLE>
<script LANGUAGE="Javascript">
<!--
function ChkField() {

    var vURealName = document.mnform.MURealName.value;
    var vSubject   = document.mnform.MSubject.value;
    var vUPassword = document.mnform.MUPassword.value;
			
    if (vURealName == "") {
        alert("이름을 입력하세요.\n");
	document.mnform.MURealName.focus();
	return false;
	}
    if (vSubject == "") {
	alert("'제목을 입력하세요.\n");
	document.mnform.MSubject.focus();
	return false;
	}
    if (vUPassword == "") {
	alert("'암호를 입력하세요.\n");
	document.mnform.MUPassword.focus();
	return false;
	}
return true;
} // end function
//    -->
</script>
</HEAD>
<BODY BGCOLOR="#ffffff">
<BR>
<CENTER>
<!-- 
  이 웹페이지의 기능과 메뉴를 보여준다. 
-->
<TABLE CELLSPACING=0 CELLPADDING=0>
<TR>
    <TD WIDTH=100 ALIGN="left">
	<STRONG>새글쓰기</STRONG>
        </TD>
    <TD WIDTH=340 ALIGN="center"> 
        &nbsp; </TD>
    <TD WIDTH=100 ALIGN="right">
        <FONT SIZE=2>
	    [<A HREF="YMList.asp?<%=CArg3%>">게시물목록</A>]
	    </FONT>
        </TD>
<TR>
</TABLE>
<!-- 
  데이터 입력 양식(Form)을 보여준다.  
-->
<TABLE BORDER=1 BORDERCOLOR="#FFFFFF" CELLSPACING=0 CELLPADDING=0>
<FORM METHOD="post" ACTION="YMNew.asp?<%=CArg3%>" NAME="mnform" 
    OnSubmit="return ChkField()">
<TR>
    <TD VALIGN="middle" BORDERCOLOR="#AAAAFF" BGCOLOR="#CCCCFF">
	<FONT COLOR="#0000FF" SIZE=2> &nbsp;이 름&nbsp; </FONT> 
        </TD>
    <TD BORDERCOLOR="#FFFFFF">
	<INPUT TYPE="text" NAME="MURealName" SIZE="15" MAXLENGTH=10>
	<FONT SIZE=2> 실명을 사용하세요. </FONT>
        </TD>
    <TD VALIGN="middle" BORDERCOLOR="#AAAAFF" BGCOLOR="#CCCCFF">
	<FONT COLOR="#ffffff" SIZE=2> &nbsp; </FONT>
	</TD>
</TR>
<TR>
    <TD VALIGN="middle" BORDERCOLOR="#AAAAFF" BGCOLOR="#CCCCFF">
        <FONT COLOR="#0000FF" SIZE=2> &nbsp;제 목&nbsp; </FONT> 
        </TD>
    <TD BORDERCOLOR="#FFFFFF">
        <INPUT TYPE="text" NAME="MSubject" SIZE="65">
        </TD>
    <TD VALIGN="middle" BORDERCOLOR="#AAAAFF" BGCOLOR="#CCCCFF">
        <FONT COLOR="#0000FF" SIZE=2> &nbsp; </FONT>
        </TD>
</TR>
<TR>
    <TD VALIGN="middle" BORDERCOLOR="#AAAAFF" BGCOLOR="#CCCCFF">
        <FONT COLOR="#0000FF" SIZE=2> &nbsp;내 용&nbsp; </FONT>
        </TD>
    <TD BORDERCOLOR="#FFFFFF">
        <TEXTAREA ROWS="10" COLS="65" NAME="MContent"></TEXTAREA>
        </TD>
    <TD VALIGN="middle" BORDERCOLOR="#AAAAFF" BGCOLOR="#CCCCFF">
        <FONT COLOR="#0000FF" SIZE=2> &nbsp; </FONT>
        </TD>
</TR>
<TR>
    <TD VALIGN="middle" BORDERCOLOR="#AAAAFF" BGCOLOR="#CCCCFF">
        <FONT COLOR="#0000FF" SIZE=2> &nbsp;암 호&nbsp; </FONT> 
        </TD>
    <TD BORDERCOLOR="#FFFFFF">
        <INPUT TYPE="text" NAME="MUPassword" SIZE="15" MAXLENGTH=15>
        <FONT SIZE=2>수정, 혹은 삭제시에 필요합니다.</FONT>
        </TD>
    <TD VALIGN="middle" BORDERCOLOR="#AAAAFF" BGCOLOR="#CCCCFF">
        <FONT COLOR="#0000FF" SIZE=2> &nbsp; </FONT>
        </TD>
</TR>
<TR>
    <TD> &nbsp; 
        </TD>
    <TD>
	<INPUT TYPE="submit" VALUE="입력완료">
	<INPUT TYPE="reset"  VALUE="새로입력">
        </TD>
    <TD> &nbsp; 
        </TD>
</TR>
</FORM>
</TABLE>
</CENTER>
</BODY>
</HTML>
