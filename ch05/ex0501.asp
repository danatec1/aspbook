<%
Jumsu = 80
%>
<HTML>
<HEAD>
<TITLE> Example : If Statements </TITLE>
</HEAD>
<BODY>
If ������ ����
<HR>
<% If Jumsu >= 60 Then %>
    �����մϴ� !!! �հ��� �Ͽ�����. <BR>
<% End If %>
<% If Jumsu >= 90 Then %>
    ����� ������ A ����Դϴ�.
<% ElseIf Jumsu >=80 Then %>
    ����� ������ B ����Դϴ�.
<% ElseIf Jumsu >=70 Then %>
    ����� ������ C ����Դϴ�.
<% ElseIf Jumsu >=60 Then %>
    ����� ������ D ����Դϴ�.
<% Else %>
    ���հ��Ͽ����ϴ�. ���� ��ȸ�� �̿��� �ּ���.
<% End IF %>   
</BODY>
</HTML>
