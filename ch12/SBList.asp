<!-- #include file="SBConn.inc" -->
<%
' 데이터베이스에서 게시물들을 읽어온다. 
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.open ConnString
sqlQuery = "SELECT * FROM Message ORDER BY MID DESC"
Set objRS = objConn.Execute(sqlQuery)
%> 
<HTML>
<HEAD> 
<TITLE> Message List </TITLE>
</HEAD>
<BODY>
<A HREF="SBNForm.asp">새글쓰기</A>
<TABLE BORDER="1">
<TR>
    <TH> 번호 </TH>
    <TH> 제목 </TH>
    <TH> 글쓴이 </TH>
</TR>
<% 
' 레코드셋에 레코드가 없으면,
IF objRS.BOF and objRS.EOF then %>
    <TR>
        <TD COLSPAN=3> 게시물이 하나도 없어요. <TD>
    </TR>
<% 
Else
' 레코드셋에 레코드들이 있으면, 한 레코드씩 반복하여 출력한다.
    Do Until objRS.EOF %>
	<TR>
	    <TD> <% = objRS("MID") %> </TD>
            <TD> <A HREF="SBView.asp?MID=<% = objRS("MID")%>">
	         <% = objRS("MSubject") %></A></TD>
            <TD> <% = objRS("MUName") %> </TD>
        </TR>
<%      objRS.MoveNext
    Loop 
End If %>
</TABLE>
</BODY>
</HTML>
<%
' 사용한 자원을 반납한다.
objRS.Close
Set objRS = Nothing
objConn.Close
Set objConn = Nothing
%>
