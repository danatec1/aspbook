<%
intN = 7
%>
<HTML>
<HEAD>
<TITLE> Example : Select Statement </TITLE>
</HEAD>
<BODY>
Select ���� ����
<HR>
<% 
Select Case intN 
    Case 1, 3, 5, 7, 9 
%>
    Ȧ���Դϴ�. 
<% Case 0, 2, 4, 6, 8 %>
    ¦���Դϴ�. 
<% Case Else %>
    �Է� �����Ͱ� ���ڸ����� �ƴմϴ�. 
<% 
End Select 
%>
</BODY>
</HTML>
