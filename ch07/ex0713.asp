<%
strVTime  = Request.Cookies("cooVTime")
strVCount = Request.Cookies("cooVCount")
' ó�� �湮�ϸ� Visit Count�� 0 ���� �����Ѵ�.
If strVCount = "" Then
    strVCount = 0
End If
intVCount = CInt(strVCount)
intVCount = intVCount + 1 
' ������ �湮 Ƚ���� �� �����Ѵ�.
Response.Cookies("cooVCount") = intVCount
' ���� ��¥�� �ð��� ���Ѵ�.
datNow = now()
' date ���� ���������� �ٲ۴�.
strNow = CStr(datNow) 
' ������ �湮 �ð��� �� �����Ѵ�.
Response.Cookies("cooVTime") = strNow
%>
<HTML>
<HEAD> 
<TITLE> Cookie : Visited Counter and Time </TITLE>
</HEAD>
<BODY>
<I>��Ű : �湮 Ƚ���� ������ �湮 �ð�</I> <BR> <BR>
����� �̰��� <% = intVCount %>��° �湮�Ͽ����ϴ�.<BR>
<%
If strVTime = "" Then %>
ù �湮�� ȯ���մϴ�.
<%
Else %>
������ �湮 �ð��� <% = strVTime %> �̾����ϴ�.
<%
End If %>
</BODY>
</HTML>
