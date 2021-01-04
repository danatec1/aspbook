<%
' Login 절차를 통과한 지를 Check 한다.
' 불법으로 액세스하면 Login 절차를 거치도록 한다. 
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
<I>회원제 어플리케이션(3rd Step)</I> <BR>
<FONT COLOR="red">성공적으로 Login 하였습니다.</FONT> <BR>
(이를 실제 어플리케이션으로 대체하면 된다.)
</BODY>
</HTML>
