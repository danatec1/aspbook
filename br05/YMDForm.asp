<!-- 
#include file="YConf.inc" 
-->
<%
' ������� �Է��� �а� �μ� ��ϵ� �����.
strMID = Request.QueryString("MID")
CArg4 = CArg3 & "&MID=" & strMID
' �ش� �Խù��� �д´�.
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
  �� ���������� ��ɰ� �޴��� �����ش�. 
-->
<TABLE CELLSPACING=0 CELLPADDING=0>
<TR>
    <TD WIDTH=100 ALIGN="left">
	<STRONG>�����ϱ�</STRONG>
        </TD>
    <TD WIDTH=200 ALIGN="center"> 
        &nbsp; </TD>
    <TD WIDTH=100 ALIGN="right">
        <FONT SIZE=2>
	    [<A HREF="YMList.asp?<%=CArg3%>">�Խù����</A>]
	    </FONT>
        </TD>
<TR>
</TABLE>
<!-- 
  ������ ���� ���(Form)�� �����ش�.  
-->
<TABLE BORDER=1 BORDERCOLOR="#FFFFFF" CELLSPACING=0 CELLPADDING=0>
<FORM METHOD="post" ACTION="YMDelete.asp?<%=CArg4%>">
<TR>
    <TD WIDTH=40 VALIGN="middle" BORDERCOLOR="#AAAAFF" BGCOLOR="#CCCCFF">
        <FONT COLOR="#0000FF" SIZE=2> &nbsp;�� ȣ&nbsp; </FONT> 
        </TD>
    <TD WIDTH=340 BORDERCOLOR="#FFFFFF">
        <INPUT TYPE="password" NAME="MUPassword" SIZE="15" MAXLENGTH=15>
        <FONT SIZE=2>������ �Էµ� ��ȣ�� �ʿ��մϴ�.</FONT>
        </TD>
    <TD WIDTH=20 VALIGN="middle" BORDERCOLOR="#AAAAFF" BGCOLOR="#CCCCFF">
        <FONT COLOR="#0000FF" SIZE=2> &nbsp; </FONT>
        </TD>
</TR>
<TR>
    <TD> &nbsp; 
        </TD>
    <TD>
	<INPUT TYPE="submit" VALUE="����">
        </TD>
    <TD> &nbsp; 
        </TD>
</TR>
</FORM>
</TABLE>
</CENTER>
</BODY>
</HTML>
