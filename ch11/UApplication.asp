<!-- #include file="Member.inc"   -->
<%
' Page : ���� ������ ��ȣ�� �μ��� �о �����Ѵ�.
strCPage = Request.QueryString("Page")
If (strCPage = "") Then
    strCPage = "1"
End If
intCPage = CInt(strCPage)
' �����ͺ��̽��� �����Ͽ� ���ڵ���� �д´�.
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
<I>ȸ�� ���</I>
<TABLE BORDER=1>
<TR>
    <TH> ���� </TH>
    <TH> ��ȣ </TH>
    <TH> E-Mail �ּ� </TH>
</TR>
<% 
' ���ڵ�¿� ���ڵ���� �ִ°� ?
If (objRS.BOF and objRS.EOF) Then 
%>
    <TR>
        <TD COLSPAN=3> ������ �ϳ��� �����. </TD>
    </TR>
<% 
' ���ڵ�¿� ���ڵ尡 ������,
Else 
'   Configuration ��� : �������� ũ�⸦ �����Ѵ�.
    objRS.PageSize = 2 
'   ���÷����� �������� ó������ �̵��Ѵ�.
    objRS.AbsolutePage = intCPage
'   ������ ���鼭 �� �������� ���ڵ���� ����Ѵ�.
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
'   ������ �޴��� �����ϴ� �κ�
    intTPage   = objRS.PageCount
'   Configuration ��� : ������ �׷� ũ�⸦ �����Ѵ�.
    intGSize = 2
'   ���� �������� �׷� ��ȣ�� �׷��� ������ ������ ��ȣ�� ���Ѵ�.
    intPGNumber = (intCPage - 1) \ intGSize
    intPGEnd = intPGNumber * intGSize
'   ���� ������ �׷� �޴� 
    If ( intPGEnd > 0 ) Then
%>
        [<A HREF="UAList.asp?Page=<%=intPGEnd%>">����</A>]
<%  End If
    LCount = intGSize
'   ���� ������ �׷��� ���� ��ȣ�� ���Ѵ�.
    intI = intPGEnd + 1
'   ���� ������ �׷��� ������ ��ȣ���� ����Ѵ�. 
    While (LCount > 0) and (intI <= intTPage)   
%>
        [<A HREF="UAList.asp?Page=<%=intI%>"><%=intI%></A>] &nbsp;
<%  
        intI = intI + 1
        LCount = LCount - 1
    Wend
'   ���� ������ �׷� �޴� 
    intI = intPGEnd + (intGSize + 1)
    If (intI <= intTPage ) Then
%>
        [<A HREF="UAList.asp?Page=<%=intI%>">����</A>]
<%  End If 
End If
objRS.Close
Set objRS = Nothing
%>
</BODY>
</HTML>
