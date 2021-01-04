<!-- #include file="Member.inc"   -->
<%
' Page : 현재 페이지 번호를 인수로 읽어서 결정한다.
strCPage = Request.QueryString("Page")
If (strCPage = "") Then
    strCPage = "1"
End If
intCPage = CInt(strCPage)
' 데이터베이스에 연결하여 레코드셋을 읽는다.
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.CursorType = 3
ConnString = "DSN=aspbook;UID=sa;PWD=;"
sqlQuery = "SELECT * FROM Account" 
objRS.Open sqlQuery, ConnString
%>
<HTML>
<HEAD>
<TITLE> User Account List </TITLE>
</HEAD>
<BODY>
<I>회원 목록</I>
<TABLE BORDER=1>
<TR>
    <TH> 계정 </TH>
    <TH> 암호 </TH>
    <TH> E-Mail 주소 </TH>
</TR>
<% 
' 레코드셋에 레코드들이 있는가 ?
If (objRS.BOF and objRS.EOF) Then 
%>
    <TR>
        <TD COLSPAN=3> 계정이 하나도 없어요. </TD>
    </TR>
<% 
' 레코드셋에 레코드가 있으면,
Else 
'   Configuration 상수 : 페이지의 크기를 지정한다.
    objRS.PageSize = 2 
'   디스플레이할 페이지의 처음으로 이동한다.
    objRS.AbsolutePage = intCPage
'   루프를 돌면서 한 페이지의 레코드들을 출력한다.
    PSize = objRS.PageSize
    While (PSize > 0) and (NOT objRS.EOF) 
%>
        <TR>
            <TD> <% = objRS("login") %> </TD>
            <TD> <% = objRS("password") %> </TD>
            <TD> <% = objRS("email") %> </TD>
        </TR>
<%
        objRS.MoveNext
	PSize = PSize -1
    Wend
%>
</TABLE>
<% 
'   페이지 메뉴를 관리하는 부분
    intTPage   = objRS.PageCount
'   Configuration 상수 : 페이지 그룹 크기를 지정한다.
    intGSize = 2
'   이전 페이지의 그룹 번호와 그룹의 마지막 페이지 번호를 구한다.
    intPGNumber = (intCPage - 1) \ intGSize
    intPGEnd = intPGNumber * intGSize
'   이전 페이지 그룹 메뉴 
    If ( intPGEnd > 0 ) Then
%>
        [<A HREF="UAList.asp?Page=<%=intPGEnd%>">이전</A>]
<%  End If
    LCount = intGSize
'   현재 페이지 그룹의 시작 번호를 구한다.
    intI = intPGEnd + 1
'   현재 페이지 그룹의 페이지 번호들을 출력한다. 
    While (LCount > 0) and (intI <= intTPage)   
%>
        [<A HREF="UAList.asp?Page=<%=intI%>"><%=intI%></A>] &nbsp;
<%  
        intI = intI + 1
        LCount = LCount - 1
    Wend
'   다음 페이지 그룹 메뉴 
    intI = intPGEnd + (intGSize + 1)
    If (intI <= intTPage ) Then
%>
        [<A HREF="UAList.asp?Page=<%=intI%>">다음</A>]
<%  End If 
End If
objRS.Close
Set objRS = Nothing
%>
</BODY>
</HTML>
