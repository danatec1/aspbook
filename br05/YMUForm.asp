<!-- 
#include file="YConf.inc" 
-->
<%
' 사용자의 입력을 읽고 인수 목록도 만든다.
strMID = Request.QueryString("MID")
CArg4 = CArg3 & "&MID=" & strMID
' 해당 게시물을 읽는다.
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
sqlQuery = "SELECT * FROM " & strTable 
sqlQuery = sqlQuery & " WHERE MID =" & strMID
Set objRS = objConn.Execute(SQLQuery)
strContent = objRS("MContent")
%>
<HTML>
<HEAD>
<TITLE> Message Update Form </TITLE>
<script LANGUAGE="Javascript">
<!--
function ChkField() {
    var vURealName = document.muform.MURealName.value;
    var vSubject   = document.muform.MSubject.value;
    var vUPassword = document.muform.MUPassword.value;
			
    if (vURealName == "") {
        alert("이름을 입력하세요.\n");
	document.muform.MURealName.focus();
	return false;
	}
    if (vSubject == "") {
	alert("'제목을 입력하세요.\n");
	document.muform.MSubject.focus();
	return false;
	}
    if (vUPassword == "") {
	alert("'암호를 입력하세요.\n");
	document.muform.MUPassword.focus();
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
	<STRONG>수정하기</STRONG>
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
  데이터 수정 양식(Form)을 보여준다.  
-->
<TABLE BORDER=1 BORDERCOLOR="#FFFFFF" CELLSPACING=0 CELLPADDING=0>
<FORM METHOD="post" ACTION="YMUpdate.asp?<%=CArg4%>" NAME="muform" 
    OnSubmit="return ChkField()">
<TR>
    <TD VALIGN="middle" BORDERCOLOR="#AAAAFF" BGCOLOR="#CCCCFF">
	<FONT COLOR="#0000FF" SIZE=2> &nbsp;이 름&nbsp; </FONT> 
        </TD>
    <TD BORDERCOLOR="#FFFFFF">
	<INPUT TYPE="text" NAME="MURealName" SIZE="15" MAXLENGTH=10
            VALUE="<%=objRS("MURealName")%>">
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
        <INPUT TYPE="text" NAME="MSubject" SIZE="65"
            VALUE="<%=objRS("MSubject")%>">
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
        <TEXTAREA ROWS="10" COLS="65" 
            NAME="MContent"><%=strContent%></TEXTAREA>
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
        <INPUT TYPE="password" NAME="MUPassword" SIZE="15" MAXLENGTH=15>
        <FONT SIZE=2>이전에 입력된 암호가 필요합니다.</FONT>
        </TD>
    <TD VALIGN="middle" BORDERCOLOR="#AAAAFF" BGCOLOR="#CCCCFF">
        <FONT COLOR="#0000FF" SIZE=2> &nbsp; </FONT>
        </TD>
</TR>
<TR>
    <TD> &nbsp; 
        </TD>
    <TD>
	<INPUT TYPE="submit" VALUE="수정완료">
	<INPUT TYPE="reset"  VALUE="처음부터">
        </TD>
    <TD> &nbsp; 
        </TD>
</TR>
</FORM>
</TABLE>
</CENTER>
</BODY>
</HTML>
