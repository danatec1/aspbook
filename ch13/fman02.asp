<% 
' FileUpLoad ��ü�� ����Ͽ� �Է� �����͸� �д´�.
Set objFUL = Server.CreateObject("FileMan.FileUpLoad")
strName = objFUL.GetValueByName("name")
strAge  = objFUL.GetValueByName("age")
%>
<HTML>
<HEAD> 
<TITLE> FileUpload Object (Result) </TITLE>
</HEAD>
<BODY>
<I>FileUpload ��ü (���)</I><BR>
����� �̸��� <% = strName %>�̰�, <BR>  
���̴� <% = strAge %>���Դϴ�.
</BODY>
</HTML>
