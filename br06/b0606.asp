<%
' 객체를 생성하고 속성들을 설정한다.
Set objMail = Server.CreateObject("CDONTS.NewMail")
objMail.To      = Request.Form("EMTo")
objMail.From    = Request.Form("EMFrom")
objMail.Subject = Request.Form("EMSubject")
objMail.Body    = Request.Form("EMBody")
' 파일을 첨부한다.
objMail.AttachFile Request.Form("EMFPath")
' 메일을 전송하고 자원을 반납한다.
objMail.Send 
Set objMail = Nothing
%>
<HTML>
<HEAD>
<TITLE> Mail Send </TITLE>
</HEAD>
<BODY>
첨부 화일이 있는 메일을 전송하였습니다.
</BODY>
</HTML>
 