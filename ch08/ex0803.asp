<%
' Login ������ ����� ���� Check �Ѵ�.
' �ҹ����� �׼����ϸ� Login ������ ��ġ���� �Ѵ�. 
strLogin = Session("sesLogin")
strPassword = Session("sesPassword")
If strLogin = "" Then
    Response.Redirect "ex0801.asp"
ElseIf strPassword = "" Then
    Response.Redirect "ex0801.asp"
End If
%>
<HTML>
<HEAD>
<TITLE> Membership Application (3rd step) </TITLE>
</HEAD>
<BODY>
<I>ȸ���� ���ø����̼�(3rd Step)</I> <BR>
<FONT COLOR="red">���������� Login �Ͽ����ϴ�.</FONT> <BR>
(�̸� ���� ���ø����̼����� ��ü�ϸ� �ȴ�.)
</BODY>
</HTML>
