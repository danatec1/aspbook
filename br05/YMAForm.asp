<!-- 
#include file="YConf.inc" 
-->
<%
' ������� �Է��� �а� �μ� ��ϵ� �����.
strMID = Request.QueryString("MID")
CArg4 = CArg3 & "&MID=" & strMID
%>
<HTML>
<HEAD>
<TITLE> Message Answer Form </TITLE>
<script LANGUAGE="Javascript">
<!--
function ChkField() {

    var vURealName = document.maform.MURealName.value;
    var vSubject   = document.maform.MSubject.value;
    var vUPassword = document.maform.MUPassword.value;
			
    if (vURealName == "") {
        alert("�̸��� �Է��ϼ���.\n");
	document.maform.MURealName.focus();
	return false;
	}
    if (vSubject == "") {
	alert("������ �Է��ϼ���.\n");
	document.maform.MSubject.focus();
	return false;
	}
    if (vUPassword == "") {
	alert("��ȣ�� �Է��ϼ���.\n");
	document.maform.MUPassword.focus();
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
	<STRONG>���侲��</STRONG>
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
  ������ �Է� ���(Form)�� �����ش�.  
-->
<TABLE BORDER=1 BORDERCOLOR="#FFFFFF" CELLSPACING=0 CELLPADDING=0>
<FORM METHOD="post" ACTION="YMAnswer.asp?<%=CArg4%>" NAME="maform" 
    OnSubmit="return ChkField()">
<TR>
    <TD VALIGN="middle" BORDERCOLOR="#AAAAFF" BGCOLOR="#CCCCFF">
	<FONT COLOR="#0000FF" SIZE=2> &nbsp;�� ��&nbsp; </FONT> 
        </TD>
    <TD BORDERCOLOR="#FFFFFF">
	<INPUT TYPE="text" NAME="MURealName" SIZE="15" MAXLENGTH=10>
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
        <INPUT TYPE="text" NAME="MSubject" SIZE="65">
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
        <TEXTAREA ROWS="10" COLS="65" NAME="MContent"></TEXTAREA>
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
        <INPUT TYPE="text" NAME="MUPassword" SIZE="15" MAXLENGTH=15>
        <FONT SIZE=2>����, Ȥ�� �����ÿ� �ʿ��մϴ�.</FONT>
        </TD>
    <TD VALIGN="middle" BORDERCOLOR="#AAAAFF" BGCOLOR="#CCCCFF">
        <FONT COLOR="#0000FF" SIZE=2> &nbsp; </FONT>
        </TD>
</TR>
<TR>
    <TD> &nbsp; 
        </TD>
    <TD>
	<INPUT TYPE="submit" VALUE="�Է¿Ϸ�">
	<INPUT TYPE="reset"  VALUE="�����Է�">
        </TD>
    <TD> &nbsp; 
        </TD>
</TR>
</FORM>
</TABLE>
</CENTER>
</BODY>
</HTML>
