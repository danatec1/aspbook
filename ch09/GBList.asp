<%
On Error Resume Next 
' ������ ������ ������ ���丮 �̸��� ã�´�.
strPath = Server.MapPath("/booksrc/ch09")
strPFile = strPath & "/" & "gbook.txt"
' FileSystemObject ��ü�� �����Ѵ�.
Set objFSO = CreateObject("Scripting.FileSystemObject")
' ���Ͽ� ������ �б�(1) ���ؼ� Open �Ѵ�.
' ����, ������ �������� ������ ������ ����(True)�Ѵ�.
Set objTS = objFSO.OpenTextFile(strPFile, 1, True)
%>
<HTML>
<HEAD>
<TITLE> Guest Book : List </TITLE>
</HEAD>
<BODY BGCOLOR=white>
<CENTER>
<H2> ���� </H2>
<A HREF="GBNForm.asp">�۾���</A>
<TABLE>
<% 
If objTS.AtEndOfStream Then %>
<TABLE BORDER=1 WIDTH=400>
    <TR>
        <TD> ���Ͽ� ���� ���� �����ϴ�. </TD>
    </TR>
</TABLE>
<%
Else 
    Do Until objTS.AtEndOfStream %>
        <TABLE BORDER=1 WIDTH=400>
        <TR>
            <TD WIDTH=50> �̸� </TD> <TD> <% = objTS.ReadLine %> </TD>
        </TR> <TR> 
            <TD WIDTH=50> ���� </TD> <TD> <% = objTS.ReadLine %> </TD>
        </TR> <TR> 
            <TD WIDTH=50> �Ͻ� </TD> <TD> <% = objTS.ReadLine %> </TD>
        </TR>
        </TABLE>
	<BR>
<%  Loop
End If
' ��ü ����� �ߴ��ϰ� �ı��Ѵ�.
objTS.Close
Set objTS = Nothing
%>
</CENTER>
</BODY>
</HTML>
