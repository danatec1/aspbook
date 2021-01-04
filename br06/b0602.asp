<%
Set objMail = Server.CreateObject("CDONTS.NewMail")

objMail.To      = Request.Form("EMTo")
objMail.From    = Request.Form("EMFrom")
objMail.Subject = Request.Form("EMSubject")
objMail.Body    = Request.Form("EMBody")

objMail.Send 
Set objMail = Nothing
%>
<HTML>
<HEAD>
<TITLE> Mail Send </TITLE>
</HEAD>
<BODY>
메일을 전송하였습니다.
</BODY>
</HTML>
 