<%
' HTML 문서의 앞 부분을 만든다.
strHTML1 = "<HTML>"
strHTML1 = strHTML1 & "<HEAD>"
strHTML1 = strHTML1 & "<TITLE> </TITLE>"
strHTML1 = strHTML1 & "</HEAD>"
strHTML1 = strHTML1 & "<BODY BGCOLOR=yellow>"
' HTML 문서의 뒷 부분을 만든다.
strHTML2 = "</BODY>"
strHTML2 = strHTML2 & "</HTML>"
' 객체를 생성하고 속성들을 설정한다.
Set objMail = Server.CreateObject("CDONTS.NewMail")
objMail.To      = Request.Form("EMTo")
objMail.From    = Request.Form("EMFrom")
objMail.Subject = Request.Form("EMSubject")
' 문서의 본체를 작성한다.
strHTML = strHTML1
strHTML = strHTML & Request.Form("EMBody")
strHTML = strHTML & strMTML2
objMail.Body    = strHTML
objMail.MailFormat = 0    ' 먼저, MIME 형식 사용으로 지정한다. 
objMail.BodyFormat = 0    ' 본체의 내용을 HTML 문서로 지정한다.
' 문서를 전송하고 사용한 자원을 반납한다.
objMail.Send 
Set objMail = Nothing
%>
<HTML>
<HEAD>
<TITLE> Mail Send </TITLE>
</HEAD>
<BODY>
HTML 형식으로 메일을 전송하였습니다.
</BODY>
</HTML>
 