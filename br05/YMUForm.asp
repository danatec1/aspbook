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
<TITLE> Message Update Form </TITLE>
<script LANGUAGE="Javascript">
<!--
function ChkField() {
    var vURealName = document.muform.MURealName.value;
    var vSubject   = document.muform.MSubject.value;
    var vUPassword = document.muform.MUPassword.value;
			
    if (vURealName == "") {
        alert("�̸��� �Է��ϼ���.\n");
	document.muform.MURealName.focus();
	return false;
	}
    if (vSubject == "") {
	alert("'������ �Է��ϼ���.\n");
	document.muform.MSubject.focus();
	return false;
	}
    if (vUPassword == "") {
	alert("'��ȣ�� �Է��ϼ���.\n");
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
  �� ���������� ��ɰ� �޴��� �����ش�. 
-->
<TABLE CELLSPACING=0 CELLPADDING=0>
<TR>
    <TD WIDTH=100 ALIGN="left">
	<STRONG>�����ϱ�</STRONG>
        </TD>
    <TD WIDTH=340 ALIGN="center"> 
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
<FORM METHOD="post" ACTION="YMUpdate.asp?<%=CArg4%>" NAME="muform" 
    OnSubmit="return ChkField()">
<TR>
    <TD VALIGN="middle" BORDERCOLOR="#AAAAFF" BGCOLOR="#CCCCFF">
	<FONT COLOR="#0000FF" SIZE=2> &nbsp;�� ��&nbsp; </FONT> 
        </TD>
    <TD BORDERCOLOR="#FFFFFF">
	<INPUT TYPE="text" NAME="MURealName" SIZE="15" MAXLENGTH=10
            VALUE="<%=objRS("MURealName")%>">
	<FONT SIZE=2> �Ǹ��� ����ϼ���. </FONT>
        </TD>
    <TD VALIGN="middle" BORDERCOLOR="#AAAAFF" BGCOLOR="#CCCCFF">
	<FONT COLOR="#ffffff" SIZE=2> &nbsp; </FONT>
	</TD>
</TR>
<TR>
    <TD VALIGN="middle" BORDERCOLOR="#AAAAFF" BGCOLOR="#CCCCFF">
        <FONT COLOR="#0000FF" SIZE=2> &nbsp;�� ��&nbsp; </FONT> 
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
        <FONT COLOR="#0000FF" SIZE=2> &nbsp;�� ��&nbsp; </FONT>
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
        <FONT COLOR="#0000FF" SIZE=2> &nbsp;�� ȣ&nbsp; </FONT> 
        </TD>
    <TD BORDERCOLOR="#FFFFFF">
        <INPUT TYPE="password" NAME="MUPassword" SIZE="15" MAXLENGTH=15>
        <FONT SIZE=2>������ �Էµ� ��ȣ�� �ʿ��մϴ�.</FONT>
        </TD>
    <TD VALIGN="middle" BORDERCOLOR="#AAAAFF" BGCOLOR="#CCCCFF">
        <FONT COLOR="#0000FF" SIZE=2> &nbsp; </FONT>
        </TD>
</TR>
<TR>
    <TD> &nbsp; 
        </TD>
    <TD>
	<INPUT TYPE="submit" VALUE="�����Ϸ�">
	<INPUT TYPE="reset"  VALUE="ó������">
        </TD>
    <TD> &nbsp; 
        </TD>
</TR>
</FORM>
</TABLE>
</CENTER>
</BODY>
</HTML>
