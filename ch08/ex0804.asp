<% 
' test�� ���� �������� ĳ������ �ʵ��� ����
Response.Expires = 0
' Timeout �Ӽ��� �����ϱ� ���ؼ� 1������ ���� 
Session.Timeout = 1
%>
<%
' ���� ������ ������ �ʱⰪ�� 0���� �Ѵ�.
If Session("sesVisited") = "" Then
    Session("sesVisited") = 0
End If
' ���� ������ ���� �ϳ� ������Ų��.
Session("sesVisited") = Session("sesVisited") + 1
' ���� ������ ���� 4�� �Ǹ� ��ü�� �ı��Ѵ�.
If (Session("sesVisited") = "5") Then
    Session.Abandon()
End If
%>
<HTML>
<HEAD> 
<TITLE> SessionID, Timeout Attribute and Abandon method </TITLE>
</HEAD>
<BODY>
<I>SessionID, Timeout Attribute and Abandon method</I> <BR>
<FONT COLOR="red">1�� �Ŀ� Session�� ����ȴ�.</FONT><BR>
SessionID = <% = Session.SessionID %> <BR>
�湮�� : <% = Session("sesVisited") %> <BR>
<A HREF="ex0804.asp">�ٽ��б�</A>
</BODY>
</HTML>
