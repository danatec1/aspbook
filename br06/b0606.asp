<%
' ��ü�� �����ϰ� �Ӽ����� �����Ѵ�.
Set objMail = Server.CreateObject("CDONTS.NewMail")
objMail.To      = Request.Form("EMTo")
objMail.From    = Request.Form("EMFrom")
objMail.Subject = Request.Form("EMSubject")
objMail.Body    = Request.Form("EMBody")
' ������ ÷���Ѵ�.
objMail.AttachFile Request.Form("EMFPath")
' ������ �����ϰ� �ڿ��� �ݳ��Ѵ�.
objMail.Send 
Set objMail = Nothing
%>
<HTML>
<HEAD>
<TITLE> Mail Send </TITLE>
</HEAD>
<BODY>
÷�� ȭ���� �ִ� ������ �����Ͽ����ϴ�.
</BODY>
</HTML>
 