<!-- 
#include file="YConf.inc" 
-->
<%
' Search �� �����մϴ�.
strSField = Request.QueryString("SField")
strArg4 = CArg3 & "&SField=" & strSField
strSString = Request.QueryString("SString")
strArg5 = strArg4 & "&SString=" & strSString
'
' �����ͺ��̽����� �Խù����� �о�´�. 
'
Set objRS = Server.CreateObject("ADODB.Recordset")
objRS.CursorType = 3
'
sqlQuery = "SELECT * FROM " & strTable 
sqlQuery = sqlQuery & " WHERE " & strSField 
sqlQuery = sqlQuery & " LIKE '%" & strSString & "%'" 
' Mode�� ���� ���ڵ���� �����ϵ��� �Ѵ�.
IF strMode = "qa" Then
    sqlQuery = sqlQuery & " ORDER BY MGrp DESC, MSeq"
Else
    sqlQuery = sqlQuery & " ORDER BY MID DESC"
End If
objRS.Open sqlQuery, ConnString
' 
' Configuration ��� : ������ ũ�⸦ �����Ѵ�.
objRS.PageSize = 10
' ��ü �Խù� ������ �� ������ ���� �˾Ƴ���. 
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
  �� ���������� ��ɰ� �޴��� �����ش�. 
-->
<TABLE CELLSPACING=0 CELLPADDING=0>
<TR>
    <TD WIDTH=100 ALIGN="left">
	    <STRONG>�˻� ���</STRONG>
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
        [<A HREF="YMSList.asp?<%=TArgs%>">����������</A>]
        </FONT>
<%  
Else  
        TArgs = CArg1 & "&Mode=qa" 
        TArgs = TArgs & "&Page=" & strCPage
	TArgs = TArgs & "&SField=" & strSField
	TArgs = TArgs & "&SString=" & strSString
%>
        <FONT SIZE=2>
        [<A HREF="YMSList.asp?<%=TArgs%>">����������</A>]
        </FONT>
<% 
End If 
%>       
        <FONT SIZE=2>
        [<A HREF="YMNForm.asp?<%=CArg3%>">���۾���</A>]
	</FONT>
        </TD>
    <TD WIDTH=140 ALIGN="right"> 
        <FONT SIZE=2>
	[<A HREF="YMList.asp?<%=CArg3%>">�Խù����</A>]
��      </FONT>
        </TD>
</TR>
</TABLE>
<!-- 
  �Խù� ����� �����ش�. 
-->
<TABLE CELLSPACING=0 CELLPADDING=2 BORDER=1 BORDERCOLOR="#ffffff">
<!--�Խ����� �� ���� �̸��� ����Ѵ�. --->
<!-- TR BORDERCOLOR="#bc7900" -->
<TR BORDERCOLOR="#AAAAFF">
    <TH WIDTH="40" BGCOLOR="#CCCCFF">
	    <FONT COLOR="#0000FF" SIZE=2>��ȣ</FONT></TH>
    <TH WIDTH="340" BGCOLOR="#CCCCFF">
	    <FONT COLOR="#0000FF" SIZE=2>����</FONT></TH>
    <TH WIDTH="40" BGCOLOR="#CCCCFF">
	    <FONT COLOR="#0000FF" SIZE=2>��ȸ</FONT></TH>
    <TH WIDTH="60" BGCOLOR="#CCCCFF">
	    <FONT COLOR="#0000FF" SIZE=2>�����</FONT></TH>
    <TH WIDTH="60" BGCOLOR="#CCCCFF">
	    <FONT COLOR="#0000FF" SIZE=2>�����</FONT></TH>
</TR>
<%
' �Խù��� �ִ� �� ���� ���� �����մϴ�.
If (objRS.BOF and objRS.EOF) Then
    Response.Write "<TR> <TD COLSPAN=5>"
    Response.Write "�Խù��� �ϳ��� �����ϴ�."
    Response.Write "</TD> </TR>"
Else
' ���÷����� �������� ó������ �̵��Ѵ�.
    objRS.AbsolutePage = intCPage   
' ����� ���ڵ带 �ִ��� ������ ���鼭 ����Ѵ�.
    RCount = objRS.PageSize
    Do While (NOT objRS.EOF) and (RCount > 0) 
%>
        <TR ALIGN="center" BGCOLOR="#EEEEEE">
	    <!-- �޼��� ��ȣ�� ����Ѵ� -->
	    <TD ALIGN="right">
	        <FONT SIZE=2> 
	        <% = objRS("MID") %> &nbsp; </FONT>
	        </TD>
	    <!-- ������ ����Ѵ�. -->
	    <TD ALIGN="left"> 
<%
'       �������̸�, ���⸦ �Ѵ�.
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
            <!-- ��ȸ���� ����Ѵ�. -->
	    <TD ALIGN="right">
	        <FONT SIZE=2> <%=objRS("MVisited")%> &nbsp; </FONT>
	        </TD>
	    <!-- �ۼ��ڸ� ����Ѵ�. -->
	    <TD ALIGN="left">
	        <FONT SIZE=2> &nbsp;<%=objRS("MURealName")%> </FONT>
	        </TD>
	    <!-- ��¥�� ����Ѵ�. -->
<%
            strWTime = objRS("MWTime")
            strWDate = Left(strWTime,8)
%>
	    <TD> 
	        <FONT SIZE=2> <%=strWDate%> </FONT>
	        </TD>
        </TR>
<%
'       Ŀ���� ���� ���ڵ�� �̵��մϴ�.
        objRS.MoveNext
        RCount = RCount -1
    Loop
End If
%>
</TABLE>
<!-- 
  ������ �޴��� �����ϴ� �κ�
-->
<TABLE>
<TR>
    <TD>
<% 
' Configuration ��� : ������ �׷� ũ�⸦ �����Ѵ�.
    intGSize = 10
' ���� �������� �׷� ��ȣ�� �׷��� ������ ������ ��ȣ�� ���Ѵ�.
    intPGNumber = (intCPage - 1) \ intGSize
    intPGEnd = intPGNumber * intGSize
' ���� ������ �޴� 
    TArgs = CArg2 & "&Page=" & intPGEnd 
    TArgs = TArgs & "&SField=" & strSField
    TArgs = TArgs & "&SString=" & strSString
    If ( intPGEnd > 0 ) Then
%>
        [<A HREF="YMSList.asp?<%=TArgs%>">����</A>]
<%  End If
' ���� ������ �׷��� �ѹ����� ���÷��� �Ѵ�. 
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
' ���� ������ �޴� 
    intI = intPGEnd + (intGSize + 1)
    TArgs = CArg2 & "&Page=" & intI 
    TArgs = TArgs & "&SField=" & strSField
    TArgs = TArgs & "&SString=" & strSString
    If (intI <= intTPage ) Then
%>
	[<A HREF="YMSList.asp?<%=TArgs%>">����</A>]
<%  
    End If 
%>
    </TD>
</TR>
</TABLE>
<!--
  �˻� �޴��� �����ִ� �κ� 
-->
<FORM METHOD="get" ACTION="YMSList.asp"> 
    <INPUT TYPE="hidden" NAME="Table" Value="<%=strTable%>">
    <INPUT TYPE="hidden" NAME="Mode" Value="<%=strMode%>">
    <SELECT NAME="SField">
        <OPTION VALUE="MURealName" SELECTED> �����
        <OPTION VALUE="MSubject"> ����
        <OPTION VALUE="MContent"> ����
        <OPTION VALUE="MWTime"> �����
        <OPTION VALUE="MID"> ��ȣ
    </SELECT> 
    <INPUT TYPE="text" NAME="SString">
    <INPUT TYPE=submit VALUE="�˻�">
</FORM>
</CENTER>
</BODY>
</HTML>
<%
objRS.Close
Set objRS = Nothing
%>
