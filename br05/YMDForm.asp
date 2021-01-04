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
<TITLE> Message Delete Form </TITLE>
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
	<STRONG>삭제하기</STRONG>
        </TD>
    <TD WIDTH=200 ALIGN="center"> 
        &nbsp; </TD>
    <TD WIDTH=100 ALIGN="right">
        <FONT SIZE=2>
	    [<A HREF="YMList.asp?<%=CArg3%>">게시물목록</A>]
	    </FONT>
        </TD>
<TR>
</TABLE>
<!-- 
  데이터 삭제 양식(Form)을 보여준다.  
-->
<TABLE BORDER=1 BORDERCOLOR="#FFFFFF" CELLSPACING=0 CELLPADDING=0>
<FORM METHOD="post" ACTION="YMDelete.asp?<%=CArg4%>">
<TR>
    <TD WIDTH=40 VALIGN="middle" BORDERCOLOR="#AAAAFF" BGCOLOR="#CCCCFF">
        <FONT COLOR="#0000FF" SIZE=2> &nbsp;암 호&nbsp; </FONT> 
        </TD>
    <TD WIDTH=340 BORDERCOLOR="#FFFFFF">
        <INPUT TYPE="password" NAME="MUPassword" SIZE="15" MAXLENGTH=15>
        <FONT SIZE=2>이전에 입력된 암호가 필요합니다.</FONT>
        </TD>
    <TD WIDTH=20 VALIGN="middle" BORDERCOLOR="#AAAAFF" BGCOLOR="#CCCCFF">
        <FONT COLOR="#0000FF" SIZE=2> &nbsp; </FONT>
        </TD>
</TR>
<TR>
    <TD> &nbsp; 
        </TD>
    <TD>
	<INPUT TYPE="submit" VALUE="삭제">
        </TD>
    <TD> &nbsp; 
        </TD>
</TR>
</FORM>
</TABLE>
</CENTER>
</BODY>
</HTML>
