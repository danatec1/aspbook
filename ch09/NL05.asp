<%
strLFile = "NL01.txt"
Set objNL = Server.CreateObject("MSWC.NextLink")
intLCount = objNL.GetListCount(strLFile)
intLIndex = objNL.GetListIndex(strLFile)
%>
<HTML>
<HEAD>
<TITLE> Windows OS </TITLE>
</HEAD>
<BODY>
<I>Windows � ü��</I> <BR>
���̷����� �ɷ��� �� ��ġ�� �����ϴ� � ü�� !!!<BR>
<!-- �̵� ��ũ ����� ���� -->
<% 
' ���� �� �������� �̵� ��ũ
If (intLIndex > 1) Then
    strPrevURL = objNL.GetPreviousURL(strLFile) %>
    [<A HREF="<%=strPrevURL%>">����</A>] &nbsp;
<% 
End If
' ���� �� �������� �̵� ��ũ
If (intLIndex < intLCount) Then
    strNextURL = objNL.GetNextURL(strLFile) %>
    [<A HREF="<%=strNextURL%>">����</A>]
<% 
End If %>
<!-- �̵� ��ũ ����� �� -->
</BODY>
</HTML>
