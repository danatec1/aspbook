<!-- 
#include file="YConf.inc" 
-->
<%
' ������� �Է��� �а� �μ� ��ϵ� �����.
strMID = Request.QueryString("MID")
CArg4 = CArg3 & "&MID=" & strMID
'
' �ش� �Խù� ��ȸ���� �ϳ� ������Ų��.
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open ConnString
sqlQuery = "UPDATE " & strTable 
sqlQuery = sqlQuery & " SET MVisited=MVisited + 1 "
sqlQuery = sqlQuery & " WHERE MID =" & strMID
objConn.Execute(SQLQuery)
'
' �ش� �Խù��� �ٽ� �д´�.
sqlQuery = "SELECT * FROM " & strTable 
sqlQuery = sqlQuery & " WHERE MID =" & strMID
Set objRS = objConn.Execute(SQLQuery)
' ��Ÿ ó��
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
  �� ���������� ��ɰ� �޴��� �����ش�. 
-->
<TABLE CELLSPACING=0 CELLPADDING=0>
<TR>
    <TD WIDTH=240 ALIGN="left">
        <STRONG>���б�</STRONG>
        </TD>
    <TD WIDTH=300 ALIGN="right">	
        <FONT SIZE=2>
	[<A HREF="YMList.asp?<%=CArg3%>">�Խù����</A>]
��      </FONT>
        </TD>
</TR>
</TABLE>
<!-- 
  �Խ����� �� ������ �����ش�. 
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
  �޴��� �����ش�.
-->
<TABLE>
<TR>
    <TD>
        <FONT SIZE=2>
	[<A HREF="YMAForm.asp?<%=CArg4%>">���侲��</A>]
	[<A HREF="YMUForm.asp?<%=CArg4%>">�����ϱ�</A>]
	[<A HREF="YMDForm.asp?<%=CArg4%>">�����ϱ�</A>] </FONT>
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
