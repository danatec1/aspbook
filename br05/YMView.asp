<!-- 
#include file="YConf.inc" 
-->
<%
' 사용자의 입력을 읽고 인수 목록도 만든다.
strMID = Request.QueryString("MID")
CArg4 = CArg3 & "&MID=" & strMID
'
' 해당 게시물 조회수를 하나 증가시킨다.
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
sqlQuery = "UPDATE " & strTable 
sqlQuery = sqlQuery & " SET MVisited=MVisited + 1 "
sqlQuery = sqlQuery & " WHERE MID =" & strMID
objConn.Execute(SQLQuery)
'
' 해당 게시물을 다시 읽는다.
sqlQuery = "SELECT * FROM " & strTable 
sqlQuery = sqlQuery & " WHERE MID =" & strMID
Set objRS = objConn.Execute(SQLQuery)
' 기타 처리
strNewLine = chr(10) & "&gt; "
strContent = objRS("MContent")
strContent = Replace(strContent, chr(10), strNewLine)
%>
<HTML>
<HEAD>
<TITLE> Message View </TITLE>
</HEAD>
<BODY BGCOLOR="#ffffff">
<BR>
<CENTER>
<!-- 
  이 웹페이지의 기능과 메뉴를 보여준다. 
-->
<TABLE CELLSPACING=0 CELLPADDING=0>
<TR>
    <TD WIDTH=240 ALIGN="left">
        <STRONG>글읽기</STRONG>
        </TD>
    <TD WIDTH=300 ALIGN="right">	
        <FONT SIZE=2>
	[<A HREF="YMList.asp?<%=CArg3%>">게시물목록</A>]
　      </FONT>
        </TD>
</TR>
</TABLE>
<!-- 
  게시판의 글 내용을 보여준다. 
-->
<TABLE WIDTH=540 CELLSPACING=0 CELLPADDING=2 BORDER=1 BORDERCOLOR="#ffffff">
<TR>
    <TD BGCOLOR="#CCCCFF" BORDERCOLOR="#AAAAFF">
        <FONT SIZE=2 COLOR="#0000FF">
        &nbsp;&nbsp; <%=objRS("MSubject")%></FONT>
        </TD>
</TR>
<TR>
    <TD BGCOLOR="#EEEEEE"> 
	<BR>
        <PRE>&gt; <%=strContent%></PRE>
        </TD>
</TR>
<TR>
    <TD BGCOLOR="#EEEEEE"> 
        <HR> 
        </TD>
</TR>
<TR>
    <TD ALIGN="left" BGCOLOR="#EEEEEE">
        <FONT SIZE=2>
        <%=objRS("MURealName")%> : <%=objRS("MWTime")%> </FONT>
        </TD>
</TR>
</TABLE>
<!--
  메뉴를 보여준다.
-->
<TABLE>
<TR>
    <TD>
        <FONT SIZE=2>
	[<A HREF="YMAForm.asp?<%=CArg4%>">응답쓰기</A>]
	[<A HREF="YMUForm.asp?<%=CArg4%>">수정하기</A>]
	[<A HREF="YMDForm.asp?<%=CArg4%>">삭제하기</A>] </FONT>
        </TD>
</TR>
</TABLE>
</CENTER>
</BODY>
</HTML>
<%
objRS.Close
Set objRS = Nothing
objConn.Close 
Set objConn = Nothing
%>
