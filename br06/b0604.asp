<%
' HTML ������ �� �κ��� �����.
strHTML1 = "<HTML>"
strHTML1 = strHTML1 & "<HEAD>"
strHTML1 = strHTML1 & "<TITLE> </TITLE>"
strHTML1 = strHTML1 & "</HEAD>"
strHTML1 = strHTML1 & "<BODY BGCOLOR=yellow>"
' HTML ������ �� �κ��� �����.
strHTML2 = "</BODY>"
strHTML2 = strHTML2 & "</HTML>"
' ��ü�� �����ϰ� �Ӽ����� �����Ѵ�.
Set objMail = Server.CreateObject("CDONTS.NewMail")
objMail.To      = Request.Form("EMTo")
objMail.From    = Request.Form("EMFrom")
objMail.Subject = Request.Form("EMSubject")
' ������ ��ü�� �ۼ��Ѵ�.
strHTML = strHTML1
strHTML = strHTML & Request.Form("EMBody")
strHTML = strHTML & strMTML2
objMail.Body    = strHTML
objMail.MailFormat = 0    ' ����, MIME ���� ������� �����Ѵ�. 
objMail.BodyFormat = 0    ' ��ü�� ������ HTML ������ �����Ѵ�.
' ������ �����ϰ� ����� �ڿ��� �ݳ��Ѵ�.
objMail.Send 
Set objMail = Nothing
%>
<HTML>
<HEAD>
<TITLE> Mail Send </TITLE>
</HEAD>
<BODY>
HTML �������� ������ �����Ͽ����ϴ�.
</BODY>
</HTML>
 