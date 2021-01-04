<!-- #include file="SJConf.inc" -->
<%
' 데이터베이스에서 자료들을 읽어온다. 
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.open ConnString
sqlQuery = "SELECT * FROM JMessage ORDER BY MID DESC"
Set objRS = objConn.Execute(sqlQuery)
%> 
<HTML>
<HEAD> 
<TITLE> JaRyo List </TITLE>
</HEAD>
<BODY>
<A HREF="SJNForm.asp">자료올리기</A>
<TABLE BORDER="1">
<TR>
    <TH> 번호 </TH>
    <TH> 파일 </TH>
    <TH> 제목 </TH>
    <TH> 글쓴이 </TH>
</TR>
<% 
' 레코드셋에 레코드가 없으면,
IF objRS.BOF and objRS.EOF then %>
    <TR>
        <TD COLSPAN=4> 자료가 하나도 없어요. <TD>
    </TR>
<% 
Else
' 레코드셋에 레코드들이 있으면, 한 레코드씩 반복하여 출력한다.
    Do Until objRS.EOF 
	strMID = objRS("MID")
        strSubject = objRS("MSubject")
        strUName = objRS("MUName")
        strUFName = objRS("MUFName")
%>
	<TR>
	    <TD> <% = strMID %> </TD>
<%      If ( strUFName <> "") Then %>
	    <TD> @ </TD>
<%      Else %>
	    <TD> &nbsp; </TD>
<%      End If %>
            <TD> <A HREF="SJView.asp?MID=<% = strMID %>">
	         <% = strSubject %></A></TD>
            <TD> <% = strUName %> </TD>
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
