<!-- 
#include file="YConf.inc" 
-->
<%
' Search 를 지원합니다.
strSField = Request.QueryString("SField")
strArg4 = CArg3 & "&SField=" & strSField
strSString = Request.QueryString("SString")
strArg5 = strArg4 & "&SString=" & strSString
'
' 데이터베이스에서 게시물들을 읽어온다. 
'
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.CursorType = 3
'
sqlQuery = "SELECT * FROM " & strTable 
sqlQuery = sqlQuery & " WHERE " & strSField 
sqlQuery = sqlQuery & " LIKE '%" & strSString & "%'" 
' Mode에 따라 레코드셋을 정렬하도록 한다.
IF strMode = "qa" Then
    sqlQuery = sqlQuery & " ORDER BY MGrp DESC, MSeq"
Else
    sqlQuery = sqlQuery & " ORDER BY MID DESC"
End If
objRS.Open sqlQuery, ConnString
' 
' Configuration 상수 : 페이지 크기를 정의한다.
objRS.PageSize = 10
' 전체 게시물 갯수와 총 페이지 수를 알아낸다. 
If (objRS.BOF and objRS.EOF) Then
    intTRecord = 0
    intTPage   = 0
Else
    intTRecord = objRS.RecordCount
    intTPage   = objRS.PageCount
End If
%>
<HTML>
<HEAD>
<TITLE> Message Search List </TITLE>
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
	    <STRONG>검색 목록</STRONG>
        </TD>
    <TD WIDTH=300 ALIGN="center">
<%  
If strMode = "qa" Then 
        TArgs = CArg1 & "&Mode=se" 
        TArgs = TArgs & "&Page=" & strCPage
	TArgs = TArgs & "&SField=" & strSField
	TArgs = TArgs & "&SString=" & strSString
%>
        <FONT SIZE=2>
        [<A HREF="YMSList.asp?<%=TArgs%>">순서형보기</A>]
        </FONT>
<%  
Else  
        TArgs = CArg1 & "&Mode=qa" 
        TArgs = TArgs & "&Page=" & strCPage
	TArgs = TArgs & "&SField=" & strSField
	TArgs = TArgs & "&SString=" & strSString
%>
        <FONT SIZE=2>
        [<A HREF="YMSList.asp?<%=TArgs%>">응답형보기</A>]
        </FONT>
<% 
End If 
%>       
        <FONT SIZE=2>
        [<A HREF="YMNForm.asp?<%=CArg3%>">새글쓰기</A>]
	</FONT>
        </TD>
    <TD WIDTH=140 ALIGN="right"> 
        <FONT SIZE=2>
	[<A HREF="YMList.asp?<%=CArg3%>">게시물목록</A>]
　      </FONT>
        </TD>
</TR>
</TABLE>
<!-- 
  게시물 목록을 보여준다. 
-->
<TABLE CELLSPACING=0 CELLPADDING=2 BORDER=1 BORDERCOLOR="#ffffff">
<!--게시판의 각 열의 이름을 출력한다. --->
<!-- TR BORDERCOLOR="#bc7900" -->
<TR BORDERCOLOR="#AAAAFF">
    <TH WIDTH="40" BGCOLOR="#CCCCFF">
	    <FONT COLOR="#0000FF" SIZE=2>번호</FONT></TH>
    <TH WIDTH="340" BGCOLOR="#CCCCFF">
	    <FONT COLOR="#0000FF" SIZE=2>제목</FONT></TH>
    <TH WIDTH="40" BGCOLOR="#CCCCFF">
	    <FONT COLOR="#0000FF" SIZE=2>조회</FONT></TH>
    <TH WIDTH="60" BGCOLOR="#CCCCFF">
	    <FONT COLOR="#0000FF" SIZE=2>등록자</FONT></TH>
    <TH WIDTH="60" BGCOLOR="#CCCCFF">
	    <FONT COLOR="#0000FF" SIZE=2>등록일</FONT></TH>
</TR>
<%
' 게시물이 있는 지 없는 지를 결정합니다.
If (objRS.BOF and objRS.EOF) Then
    Response.Write "<TR> <TD COLSPAN=5>"
    Response.Write "게시물이 하나도 없읍니다."
    Response.Write "</TD> </TR>"
Else
' 디스플레이할 페이지의 처음으로 이동한다.
    objRS.AbsolutePage = intCPage   
' 출력할 레코드를 있는한 루프를 돌면서 출력한다.
    RCount = objRS.PageSize
    Do While (NOT objRS.EOF) and (RCount > 0) 
%>
        <TR ALIGN="center" BGCOLOR="#EEEEEE">
	    <!-- 메세지 번호를 출력한다 -->
	    <TD ALIGN="right">
	        <FONT SIZE=2> 
	        <% = objRS("MID") %> &nbsp; </FONT>
	        </TD>
	    <!-- 제목을 출력한다. -->
	    <TD ALIGN="left"> 
<%
'       응답형이면, 들어쓰기를 한다.
        If strMode = "qa" Then
            For I=1 To objRS("MLvl")
	        Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;")
            Next
        End If
%>
	        <A HREF="YMView.asp?<%=CArg3%>&MID=<%=objRS("MID")%>"> 
	        <FONT SIZE=2> 
	        <% = objRS("MSubject")%></FONT></A>
	        </TD>
            <!-- 조회수를 출력한다. -->
	    <TD ALIGN="right">
	        <FONT SIZE=2> <%=objRS("MVisited")%> &nbsp; </FONT>
	        </TD>
	    <!-- 작성자를 출력한다. -->
	    <TD ALIGN="left">
	        <FONT SIZE=2> &nbsp;<%=objRS("MURealName")%> </FONT>
	        </TD>
	    <!-- 날짜를 출력한다. -->
<%
            strWTime = objRS("MWTime")
            strWDate = Left(strWTime,8)
%>
	    <TD> 
	        <FONT SIZE=2> <%=strWDate%> </FONT>
	        </TD>
        </TR>
<%
'       커서를 다음 레코드로 이동합니다.
        objRS.MoveNext
        RCount = RCount -1
    Loop
End If
%>
</TABLE>
<!-- 
  페이지 메뉴를 관리하는 부분
-->
<TABLE>
<TR>
    <TD>
<% 
' Configuration 상수 : 페이지 그룹 크기를 지정한다.
    intGSize = 10
' 이전 페이지의 그룹 번호와 그룹의 마지막 페이지 번호를 구한다.
    intPGNumber = (intCPage - 1) \ intGSize
    intPGEnd = intPGNumber * intGSize
' 이전 페이지 메뉴 
    TArgs = CArg2 & "&Page=" & intPGEnd 
    TArgs = TArgs & "&SField=" & strSField
    TArgs = TArgs & "&SString=" & strSString
    If ( intPGEnd > 0 ) Then
%>
        [<A HREF="YMSList.asp?<%=TArgs%>">이전</A>]
<%  End If
' 현재 페이지 그룹의 넘버들을 디스플레이 한다. 
    LCount = intGSize
    intI = intPGEnd + 1
    Do While (LCount > 0) and (intI <= intTPage)   
        TArgs = CArg2 & "&Page=" & intI 
        TArgs = TArgs & "&SField=" & strSField
        TArgs = TArgs & "&SString=" & strSString
%>
	[<A HREF="YMSList.asp?<%=TArgs%>"><%=intI%></A>] 
        &nbsp;
<%  
        intI = intI + 1
        LCount = LCount - 1
    Loop
' 다음 페이지 메뉴 
    intI = intPGEnd + (intGSize + 1)
    TArgs = CArg2 & "&Page=" & intI 
    TArgs = TArgs & "&SField=" & strSField
    TArgs = TArgs & "&SString=" & strSString
    If (intI <= intTPage ) Then
%>
	[<A HREF="YMSList.asp?<%=TArgs%>">다음</A>]
<%  
    End If 
%>
    </TD>
</TR>
</TABLE>
<!--
  검색 메뉴를 보여주는 부분 
-->
<FORM METHOD="get" ACTION="YMSList.asp"> 
    <INPUT TYPE="hidden" NAME="Table" Value="<%=strTable%>">
    <INPUT TYPE="hidden" NAME="Mode" Value="<%=strMode%>">
    <SELECT NAME="SField">
        <OPTION VALUE="MURealName" SELECTED> 등록자
        <OPTION VALUE="MSubject"> 제목
        <OPTION VALUE="MContent"> 내용
        <OPTION VALUE="MWTime"> 등록일
        <OPTION VALUE="MID"> 번호
    </SELECT> 
    <INPUT TYPE="text" NAME="SString">
    <INPUT TYPE=submit VALUE="검색">
</FORM>
</CENTER>
</BODY>
</HTML>
<%
objRS.Close
Set objRS = Nothing
%>
